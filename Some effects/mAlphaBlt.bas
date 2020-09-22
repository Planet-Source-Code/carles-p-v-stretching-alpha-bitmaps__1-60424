Attribute VB_Name = "mAlphaBlt"
'================================================
' Module:        mAlphaBlt.bas
' Author:        Carles P.V.
'                (see resizing routines credits)
' Dependencies:
' Last revision: 2005.05.08
'================================================
'
' History:
'
' - 2005.04.01: First release
'
' - 2005.05.03: Speed up: checked special alpha values
'               (full opaque and full transparent)
'
' - 2005.05.08: AlphaBltStretch and AlphaBlendStretch variations:
'
'               - 'Bilinear resize' original routine from 'Reconstructor' by Peter Scale
'                 (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46515&lngWId=1)
'
'               - 'Integer maths' version from 'RVTVBIMG v2 - Image Processing in VB' by Ron van Tilburg
'                 (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=47445&lngWId=1)
'
'               Slight modifications:
'
'               - Reduced to 32-bit case.
'
'               - X and Y axes scaling LUTs.
'
'               If someone wants to use these AlphaBlendXXX (interpolated) functions in a "multi-layer"
'               application, will finish up with undesired results. These results can be appreciated if
'               you set iterations at, for example, 10. We finish up with really darken edge-pixels.
'               A correct interpolated resizing of alpha-bitmaps is quite more complex.
'               Problems come from edge blended pixels (pre-blending (interpolation) null alpha pixels
'               color information...)

Option Explicit

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds     As SAFEARRAYBOUND
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7
Private Const COLORONCOLOR = 3

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long

'//

Private Const SCALER As Long = 7983360 ' 2^8.3^4.5.7.11 ~= 2^22.928

Private Type uBiLU
    ip As Long ' integer position (integer part of number)
    sf As Long ' scaled fraction part of number (integer + fraction = exact)
    cf As Long ' ones complement of fraction
End Type

Public Function AlphaBlt( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal BackColor As OLE_COLOR, _
                ByVal hBitmap As Long _
                ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  
  Dim R         As Long
  Dim G         As Long
  Dim B         As Long
  Dim a1        As Long
  Dim a2        As Long
  
  Dim uSA       As SAFEARRAY1D
  Dim aBits()   As Byte
  
  Dim i         As Long
  Dim iIn       As Long
    
    '-- Check type (bitmap)
    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then
        
        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then
            
            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then
            
                With uBIH
                
                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = uBI.bmWidth
                    .biHeight = uBI.bmHeight
                    .biSizeImage = (4 * .biWidth) * .biHeight
                        
                    '-- Get source (image) color data
                    ReDim aBits(.biSizeImage - 1)
                    Call CopyMemory(aBits(0), ByVal uBI.bmBits, .biSizeImage)
                    
                    '-- Translate OLE color
                    Call OleTranslateColor(BackColor, 0, BackColor)
                    R = (BackColor And &HFF&)
                    G = (BackColor And &HFF00&) \ &H100
                    B = (BackColor And &HFF0000) \ &H10000
                    
                    '-- Blend with BackColor
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aBits(i)
                        If (a1 = &HFF) Then
                            '-- Do nothing
                        ElseIf (a1 = &H0) Then
                            '-- Dest. = Source (solid background)
                            aBits(i - 1) = R
                            aBits(i - 2) = G
                            aBits(i - 3) = B
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * R) \ &HFF: iIn = iIn - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * G) \ &HFF: iIn = iIn - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * B) \ &HFF
                        End If
                    Next i
                    
                    '-- Paint alpha-blended
                    AlphaBlt = StretchDIBits(hDC, x, y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, aBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
                End With
            End If
        End If
    End If
End Function

Public Function AlphaBltStretch( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByVal BackColor As OLE_COLOR, _
                ByVal Interpolate As Boolean, _
                ByVal hBitmap As Long _
                ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  Dim lPrvMode  As Long
  
  Dim R         As Long
  Dim G         As Long
  Dim B         As Long
  Dim a1        As Long
  Dim a2        As Long
  
  Dim bResize   As Boolean
  Dim bResample As Boolean
  Dim uSA       As SAFEARRAY1D
  Dim aBits()   As Byte
  Dim aBitsR()  As Byte
  
  Dim i         As Long
  Dim iIn       As Long
    
    '-- Check type (bitmap)
    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then
        
        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then
            
            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then
            
                With uBIH
                    
                    '-- Need to resample
                    bResize = (uBI.bmWidth <> Width And uBI.bmHeight <> Height)
                    bResample = (bResize And Interpolate)
                    
                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    If (bResample) Then
                        .biWidth = Width
                        .biHeight = Height
                      Else
                        .biWidth = uBI.bmWidth
                        .biHeight = uBI.bmHeight
                    End If
                    .biSizeImage = (4 * .biWidth) * .biHeight
                        
                    '-- Get source color data
                    If (bResample) Then
                        '-- Map and resize
                        Call pvMapDIBits(uSA, aBits(), ByVal uBI.bmBits, uBI.bmWidthBytes * uBI.bmHeight)
                        aBitsR() = pvBilinearResize(aBits(), uBI.bmWidth, uBI.bmHeight, .biWidth, .biHeight)
                      Else
                        '-- Get a copy and resize via GDI
                        ReDim aBitsR(.biSizeImage - 1)
                        Call CopyMemory(aBitsR(0), ByVal uBI.bmBits, .biSizeImage)
                    End If
                    
                    '-- Translate OLE color
                    Call OleTranslateColor(BackColor, 0, BackColor)
                    R = (BackColor And &HFF&)
                    G = (BackColor And &HFF00&) \ &H100
                    B = (BackColor And &HFF0000) \ &H10000
                    
                    '-- Blend with BackColor
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aBitsR(i)
                        If (a1 = &HFF) Then
                            '-- Do nothing
                        ElseIf (a1 = &H0) Then
                            '-- Dest. = Source (solid background)
                            aBitsR(i - 1) = R
                            aBitsR(i - 2) = G
                            aBitsR(i - 3) = B
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aBitsR(iIn) = (a1 * aBitsR(iIn) + a2 * R) \ &HFF: iIn = iIn - 1
                            aBitsR(iIn) = (a1 * aBitsR(iIn) + a2 * G) \ &HFF: iIn = iIn - 1
                            aBitsR(iIn) = (a1 * aBitsR(iIn) + a2 * B) \ &HFF
                        End If
                    Next i
                    
                    If (bResample) Then
                        Call pvUnmapDIBits(aBits())
                    End If
                    
                    '-- Paint alpha-blended (stretched)
                    If (bResize And Not bResample) Then
                        lPrvMode = SetStretchBltMode(hDC, COLORONCOLOR)
                        AlphaBltStretch = StretchDIBits(hDC, x, y, Width, Height, 0, 0, .biWidth, .biHeight, aBitsR(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
                        Call SetStretchBltMode(hDC, lPrvMode)
                      Else
                        AlphaBltStretch = StretchDIBits(hDC, x, y, Width, Height, 0, 0, .biWidth, .biHeight, aBitsR(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
                    End If
                End With
            End If
        End If
    End If
End Function

Public Function AlphaBlend( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal hBitmap As Long _
                ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  
  Dim lhDC      As Long
  Dim lhDIB     As Long
  Dim lhDIBOld  As Long
  
  Dim a1        As Long
  Dim a2        As Long
  
  Dim uSSA      As SAFEARRAY1D
  Dim aSBits()  As Byte
  Dim uDSA      As SAFEARRAY1D
  Dim aDBits()  As Byte
  Dim lpData    As Long
  
  Dim i         As Long
  Dim iIn       As Long
    
    '-- Check type (bitmap)
    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then
        
        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then
        
            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then
            
                With uBIH
                
                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = uBI.bmWidth
                    .biHeight = uBI.bmHeight
                    .biSizeImage = (4 * .biWidth) * .biHeight
                    
                    '-- Create a temporary DIB section, select into a DC, and
                    '   bitblt destination DC area
                    lhDC = CreateCompatibleDC(0)
                    lhDIB = CreateDIBSection(lhDC, uBIH, DIB_RGB_COLORS, lpData, 0, 0)
                    lhDIBOld = SelectObject(lhDC, lhDIB)
                    Call BitBlt(lhDC, 0, 0, uBI.bmWidth, uBI.bmHeight, hDC, x, y, vbSrcCopy)
                    
                    '-- Map destination color data
                    Call pvMapDIBits(uDSA, aDBits(), lpData, .biSizeImage)
                    
                    '-- Map source color data
                    Call pvMapDIBits(uSSA, aSBits(), uBI.bmBits, .biSizeImage)
                    
                    '-- Blend with destination
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aSBits(i)
                        If (a1 = &HFF) Then
                            '-- Dest. = Source
                            iIn = i - 1
                            aDBits(iIn) = aSBits(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBits(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBits(iIn)
                        ElseIf (a1 = &H0) Then
                            '-- Do nothing (dest. preserved)
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF
                        End If
                    Next i
                    
                    '-- Paint alpha-blended
                    AlphaBlend = StretchDIBits(hDC, x, y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, ByVal lpData, uBIH, DIB_RGB_COLORS, vbSrcCopy)
                End With
                
                '-- Unmap
                Call pvUnmapDIBits(aDBits())
                Call pvUnmapDIBits(aSBits())
                
                '-- Clean up
                Call SelectObject(lhDC, lhDIBOld)
                Call DeleteObject(lhDIB)
                Call DeleteDC(lhDC)
            End If
        End If
    End If
End Function

Public Function AlphaBlendStretch( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByVal Interpolate As Boolean, _
                ByVal hBitmap As Long _
                ) As Long

  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER

  Dim lhDC      As Long
  Dim lhDIB     As Long
  Dim lhDIBOld  As Long

  Dim a1        As Long
  Dim a2        As Long

  Dim uSSA      As SAFEARRAY1D
  Dim aSBits()  As Byte
  Dim aSBitsR() As Byte
  Dim uDSA      As SAFEARRAY1D
  Dim aDBits()  As Byte
  Dim lpData    As Long

  Dim i         As Long
  Dim iIn       As Long

  '-- Check type (bitmap)

    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then

        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then

            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then

                With uBIH

                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = Width
                    .biHeight = Height
                    .biSizeImage = (4 * .biWidth) * .biHeight

                    '-- Create a temporary DIB section, select into a DC, and
                    '   bitblt destination DC area
                    lhDC = CreateCompatibleDC(0)
                    lhDIB = CreateDIBSection(lhDC, uBIH, DIB_RGB_COLORS, lpData, 0, 0)
                    lhDIBOld = SelectObject(lhDC, lhDIB)
                    Call BitBlt(lhDC, 0, 0, Width, Height, hDC, x, y, vbSrcCopy)

                    '-- Map destination color data
                    Call pvMapDIBits(uDSA, aDBits(), lpData, .biSizeImage)

                    '-- Map source color data
                    Call pvMapDIBits(uSSA, aSBits(), uBI.bmBits, uBI.bmWidthBytes * uBI.bmHeight)
                    
                    '-- Resize source color data
                    If (uBI.bmWidth <> Width And uBI.bmHeight <> Height) Then
                        If (Interpolate) Then
                            aSBitsR() = pvBilinearResize(aSBits(), uBI.bmWidth, uBI.bmHeight, .biWidth, .biHeight)
                          Else
                            aSBitsR() = pvResize(aSBits(), uBI.bmWidth, uBI.bmHeight, .biWidth, .biHeight)
                        End If
                      Else
                        aSBitsR() = aSBits()
                    End If
                    
                    '-- Blend with destination
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aSBitsR(i)
                        If (a1 = &HFF) Then
                            '-- Dest. = Source
                            iIn = i - 1
                            aDBits(iIn) = aSBitsR(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBitsR(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBitsR(iIn)
                        ElseIf (a1 = &H0) Then
                            '-- Do nothing (dest. preserved)
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aDBits(iIn) = (a1 * aSBitsR(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBitsR(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBitsR(iIn) + a2 * aDBits(iIn)) \ &HFF
                        End If
                    Next i

                    '-- Paint alpha-blended (stretched)
                    AlphaBlendStretch = StretchDIBits(hDC, x, y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, ByVal lpData, uBIH, DIB_RGB_COLORS, vbSrcCopy)
                End With

                '-- Unmap
                Call pvUnmapDIBits(aDBits())
                Call pvUnmapDIBits(aSBits())

                '-- Clean up
                Call SelectObject(lhDC, lhDIBOld)
                Call DeleteObject(lhDIB)
                Call DeleteDC(lhDC)
            End If
        End If
    End If
End Function

'//

Private Function pvBilinearResize(ByRef aOldBits() As Byte, _
                                  ByVal OldWidth As Long, ByVal OldHeight As Long, _
                                  ByVal NewWidth As Long, ByVal NewHeight As Long _
                                  ) As Byte()
                                  
' New dimensions and old dimension must be <= 2^11.5 ie. <= 2048
  
  Dim aNewBits() As Byte
  Dim i          As Long
  
  Dim kX As Double, kY As Double
  Dim ep As Double
  
  Dim fX As Long, fY As Long, cX As Long, cY As Long
  Dim xo As Long, yo As Long, po As Long, qo As Long, ro As Long
  Dim xn As Long, yn As Long, pn As Long, qn As Long
  
  Dim OldScan As Long
  Dim NewScan As Long
  Dim Skip    As Long
    
    ' Scan lines / pixel width
    OldScan = 4 * OldWidth
    NewScan = 4 * NewWidth
    Skip = 4
    
    ' Resized 'bits' array
    ReDim aNewBits(0 To NewScan * NewHeight - 1)
    
    ' Scaled fractions
    kX = CDbl(OldWidth - 1) / CDbl(NewWidth - 1)
    kY = CDbl(OldHeight - 1) / CDbl(NewHeight - 1)
    
    ' Scaling LUTs
    ReDim yLU(NewHeight - 1) As uBiLU
    ReDim xLU(NewWidth - 1) As uBiLU
    For i = 0 To NewHeight - 1
        ep = i * kY 'exact position
        With yLU(i)
            .ip = Int(ep)
            .sf = (ep - .ip) * SCALER
            .cf = SCALER - .sf
        End With
    Next i
    For i = 0 To NewWidth - 1
        ep = i * kX 'exact position
        With xLU(i)
            .ip = Int(ep)
            .sf = (ep - .ip) * SCALER
            .cf = SCALER - .sf
        End With
    Next i
    
    pn = (NewHeight - 1) * NewScan
    
    For yn = NewHeight - 1 To 0 Step -1
        
        With yLU(yn)
            fY = .sf
            cY = .cf
            po = .ip * OldScan
        End With
        qn = pn

        For xn = 0 To NewWidth - 1
            
            With xLU(xn)
                fX = .sf
                cX = .cf
                qo = po + Skip * .ip
            End With
            ro = qo + OldScan

            If (fX = 0) Then
                
                If (fY = 0) Then
                    
                    ' integer rescale in X and Y
                    
                    aNewBits(qn) = aOldBits(qo)
                    qn = qn + 1
                    qo = qo + 1
                    aNewBits(qn) = aOldBits(qo)
                    qn = qn + 1
                    qo = qo + 1
                    aNewBits(qn) = aOldBits(qo)
                    qn = qn + 1
                    qo = qo + 1
                    aNewBits(qn) = aOldBits(qo)
                    qn = qn + 1
                    qo = qo + 1
                  
                  Else
                    
                    ' integer rescale in X and interpolate Y
                    
                    aNewBits(qn) = (cY * aOldBits(qo) + fY * aOldBits(ro)) \ SCALER
                    qn = qn + 1
                    qo = qo + 1
                    ro = ro + 1
                    aNewBits(qn) = (cY * aOldBits(qo) + fY * aOldBits(ro)) \ SCALER
                    qn = qn + 1
                    qo = qo + 1
                    ro = ro + 1
                    aNewBits(qn) = (cY * aOldBits(qo) + fY * aOldBits(ro)) \ SCALER
                    qn = qn + 1
                    qo = qo + 1
                    ro = ro + 1
                    aNewBits(qn) = (cY * aOldBits(qo) + fY * aOldBits(ro)) \ SCALER
                    qn = qn + 1
                    qo = qo + 1
                    ro = ro + 1
                End If
            
            ElseIf (fY = 0) Then
                
                ' integer rescale in Y and interpolate X
                
                aNewBits(qn) = (cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                aNewBits(qn) = (cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                aNewBits(qn) = (cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                aNewBits(qn) = (cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER
                qn = qn + 1
                qo = qo + 1
              Else
                
                ' interpolation in X and Y:
                ' Apply this formula: (1-frac) * RGB1 + frac * RGB2,
                ' where frac = fraction part of number [0;1), RGB1, RGB2 = red, green or blue part of color 1, 2
                ' It is applied 3 times for every part of color (2 times on X-axes and 1 time on Y-axes)
                ' The filter computes 1 point from 4 (2x2) surrounding points
                
                aNewBits(qn) = (cY * ((cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER) _
                              + fY * ((cX * aOldBits(ro) + fX * aOldBits(ro + Skip)) \ SCALER)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                ro = ro + 1
                aNewBits(qn) = (cY * ((cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER) _
                              + fY * ((cX * aOldBits(ro) + fX * aOldBits(ro + Skip)) \ SCALER)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                ro = ro + 1
                aNewBits(qn) = (cY * ((cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER) _
                              + fY * ((cX * aOldBits(ro) + fX * aOldBits(ro + Skip)) \ SCALER)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                ro = ro + 1
                aNewBits(qn) = (cY * ((cX * aOldBits(qo) + fX * aOldBits(qo + Skip)) \ SCALER) _
                              + fY * ((cX * aOldBits(ro) + fX * aOldBits(ro + Skip)) \ SCALER)) \ SCALER
                qn = qn + 1
                qo = qo + 1
                ro = ro + 1
            End If
        Next xn
        pn = pn - NewScan
    Next yn
    
    pvBilinearResize = aNewBits
End Function

Public Function pvResize(ByRef aOldBits() As Byte, _
                         ByVal OldWidth As Long, ByVal OldHeight As Long, _
                         ByVal NewWidth As Long, ByVal NewHeight As Long _
                         ) As Byte()
                            
'Note: Slight difference when resizing (nearest) via GDI and via native routine ('rounding' issues).
'      This can be observed when scale factor is not integer.
  
  Dim aNewBits() As Byte
  Dim i          As Long
  
  Dim kX As Double, kY As Double
  Dim ep As Single
  
  Dim po As Long, qo As Long
  Dim xn As Long, yn As Long, pn As Long, qn As Long
  
  Dim OldScan As Long
  Dim NewScan As Long
  Dim Skip    As Long
    
    ' Scan lines / pixel width
    OldScan = 4 * OldWidth
    NewScan = 4 * NewWidth
    Skip = 4
    
    ' Resized 'bits' array
    ReDim aNewBits(0 To NewScan * NewHeight - 1)
    
    ' Scaled fractions
    kX = CDbl(OldWidth) / CDbl(NewWidth)
    kY = CDbl(OldHeight) / CDbl(NewHeight)
    
    ' Scaling LUTs
    ReDim yLU(NewHeight - 1) As Long
    ReDim xLU(NewWidth - 1) As Long
    For i = 0 To NewHeight - 1
        yLU(i) = Int(i * kY) * OldScan
    Next i
    For i = 0 To NewWidth - 1
        xLU(i) = Int(i * kX) * Skip
    Next i
    
    pn = (NewHeight - 1) * NewScan
    
    For yn = NewHeight - 1 To 0 Step -1
        po = yLU(yn)
        qn = pn
        For xn = 0 To NewWidth - 1
            qo = po + xLU(xn)
            ' nearest
            aNewBits(qn) = aOldBits(qo): qn = qn + 1: qo = qo + 1
            aNewBits(qn) = aOldBits(qo): qn = qn + 1: qo = qo + 1
            aNewBits(qn) = aOldBits(qo): qn = qn + 1: qo = qo + 1
            aNewBits(qn) = aOldBits(qo): qn = qn + 1: qo = qo + 1
        Next xn
        pn = pn - NewScan
    Next yn
  
    pvResize = aNewBits
End Function

Private Sub pvMapDIBits(uSA As SAFEARRAY1D, aBits() As Byte, ByVal lpData As Long, ByVal lSize As Long)
    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIBits(aBits() As Byte)
    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
