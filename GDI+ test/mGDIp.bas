Attribute VB_Name = "mGDIp"
'================================================
' Basic GDI+ testing.
' Stretching alpha-bitmaps using
' interpolation (bilinear).
'================================================
' From great stuff by Avery:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1

Option Explicit

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Enum GpStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

'//

Public Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Public Enum ImageLockMode
    [ImageLockModeRead] = &H1
    [ImageLockModeWrite] = &H2
    [ImageLockModeUserInputBuf] = &H4
End Enum

Public Enum ColorMatrixFlags
    [ColorMatrixFlagsDefault] = 0
    [ColorMatrixFlagsSkipGrays]
    [ColorMatrixFlagsAltGray]
End Enum

Public Enum ColorAdjustType
    [ColorAdjustTypeDefault] = 0
    [ColorAdjustTypeBitmap]
    [ColorAdjustTypeBrush]
    [ColorAdjustTypePen]
    [ColorAdjustTypeText]
    [ColorAdjustTypeCount]
    [ColorAdjustTypeAny]
End Enum

Public Enum InterpolationMode
    [InterpolationModeInvalid] = -1
    [InterpolationModeDefault]
    [InterpolationModeLowQuality]
    [InterpolationModeHighQuality]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Public Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Public Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault]
    [QualityModeLow]
    [QualityModeHigh]
End Enum

Public Enum CompositingMode
    [CompositingModeSourceOver] = 0
    [CompositingModeSourceCopy]
End Enum

Public Type BITMAPDATA
    Width       As Long
    Height      As Long
    Stride      As Long
    PixelFormat As Long
    Scan0       As Long
    Reserved    As Long
End Type

Public Type COLORMATRIX
    m(0 To 4, 0 To 4) As Single
End Type

Public Type ColorMap
    oldColor As Long
    newColor As Long
End Type

Public Type RECTL
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Public Const PixelFormat32bppARGB As Long = &H26200A

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus

Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, hBitmap As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus

Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As InterpolationMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As PixelOffsetMode) As GpStatus
Public Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As CompositingMode) As GpStatus
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (hAttributes As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hAttributes As Long, ByVal ColorAdjust As ColorAdjustType, ByVal EnableFlag As Long, Matrix As COLORMATRIX, GrayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hAttributes As Long) As GpStatus

Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, rect As RECTL, ByVal Flags As Long, ByVal PixelFormat As Long, LockedBitmapData As BITMAPDATA) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, LockedBitmapData As BITMAPDATA) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus

'//

Public Function StretchDIB32(oDIB As cDIB32, _
                             ByVal hDC As Long, _
                             ByVal x As Long, ByVal y As Long, _
                             ByVal nWidth As Long, ByVal nHeight As Long, _
                             Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                             Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                             Optional ByVal GlobalAlpha As Byte = 255, _
                             Optional ByVal BlendWithBackground As Boolean = True, _
                             Optional ByVal Interpolate As Boolean = False _
                             ) As Long

  Dim gplRet As Long
  
  Dim hGraphics   As Long
  Dim hAttributes As Long
  Dim uMatrix     As COLORMATRIX
  Dim hBitmap     As Long
  Dim bmpRect     As RECTL
  Dim bmpData     As BITMAPDATA
  
    If (oDIB.hDIB) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = oDIB.Width
        If (nSrcHeight = 0) Then nSrcHeight = oDIB.Height
      
        With bmpRect
            .W = oDIB.Width
            .H = oDIB.Height
        End With
        
        With bmpData
            .Width = oDIB.Width
            .Height = oDIB.Height
            .Stride = -oDIB.BytesPerScanline
            .PixelFormat = [PixelFormat32bppARGB]
            .Scan0 = oDIB.lpBits - .Stride * (oDIB.Height - 1) ' Vertical flip
        End With
        
        '-- Initialize Graphics object
        gplRet = GdipCreateFromHDC(hDC, hGraphics)
        
        '-- Initialize blank Bitmap and assign GDI DIB data
        gplRet = GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hBitmap)
        gplRet = GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeWrite] Or [ImageLockModeUserInputBuf], [PixelFormat32bppARGB], bmpData)
        gplRet = GdipBitmapUnlockBits(hBitmap, bmpData)

        '-- Prepare/Set image attributes (global alpha)
        With uMatrix
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(3, 3) = GlobalAlpha / 255
            .m(4, 4) = 1
        End With
        gplRet = GdipCreateImageAttributes(hAttributes)
        gplRet = GdipSetImageAttributesColorMatrix(hAttributes, [ColorAdjustTypeDefault], True, uMatrix, ByVal 0, [ColorMatrixFlagsDefault])
        
        '-- Draw ARGB
        gplRet = GdipSetCompositingMode(hGraphics, [CompositingModeSourceOver] * -(Not BlendWithBackground))
        gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor] + -(2 * Interpolate))
        gplRet = GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
        gplRet = GdipDrawImageRectRectI(hGraphics, hBitmap, x, y, nWidth, nHeight, xSrc, ySrc, nSrcWidth, nSrcHeight, [UnitPixel], hAttributes)
        
        '-- Clean up
        gplRet = GdipDeleteGraphics(hGraphics)
        gplRet = GdipDisposeImage(hBitmap)
        gplRet = GdipDisposeImageAttributes(hAttributes)
        
        '-- Success (not exactly)
        StretchDIB32 = (gplRet = [OK])
    End If
End Function
