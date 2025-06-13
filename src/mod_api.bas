Attribute VB_Name = "mod_api"
Option Explicit

#If Win64 Then

Rem GDI
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (Picture As t_description, RefIID As t_guid, ByVal CompletionStatus As LongPtr, image_interface As IPicture) As Long
    Public Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
    Public Declare PtrSafe Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    Public Declare PtrSafe Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
    
Rem GDI+
    Public Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As t_gdiplus, Optional ByVal outputbuf As LongPtr = 0) As LongPtr
    Public Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As LongPtr, Bitmap As LongPtr) As LongPtr
    Public Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal Bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As LongPtr
    Public Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As LongPtr
    Public Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr) As LongPtr
    
Rem Kernel
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

 #Else
 
 Rem GDI
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (Picture As typePic, RefIID As GUID, ByVal CompletionStatus As Long, image_interface As IPicture) As Long
    Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
    Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
    
Rem GDI+
    Public Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Public Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
    Public Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
    Public Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
    Public Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
    
Rem Kernel
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#End If

