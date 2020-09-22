Attribute VB_Name = "mDIB32Ext"
Option Explicit

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

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long

'//

Public Sub ConvertToGrey(oDIB As cDIB32)

  Dim uSA     As SAFEARRAY1D
  Dim aBits() As Byte
  Dim aG      As Byte
  Dim i       As Long
  
    With oDIB
        If (.hDIB) Then
            '-- Map DIB bits
            Call pvMapDIBits(uSA, aBits(), .lpBits, .Size)
            '-- Convert to grey
            For i = 3 To .Size - 1 Step 4
                If (aBits(i)) Then  'is not transparent?
                    aG = (299& * aBits(i - 1) + 587& * aBits(i - 2) + 114& * aBits(i - 3)) \ 1000
                    aBits(i - 1) = aG
                    aBits(i - 2) = aG
                    aBits(i - 3) = aG
                End If
            Next i
            '-- Unmap DIB bits
            Call pvUnmapDIBits(aBits())
        End If
    End With
End Sub

Public Sub BlendWithColor(oDIB As cDIB32, ByVal Color As OLE_COLOR, ByVal Amount As Byte)

  Dim uSA     As SAFEARRAY1D
  Dim aBits() As Byte
  Dim R       As Long
  Dim G       As Long
  Dim B       As Long
  Dim a1      As Long
  Dim a2      As Long
  Dim i       As Long
  Dim iIn     As Long
  
    With oDIB
        If (.hDIB) Then
            '-- Map DIB bits
            Call pvMapDIBits(uSA, aBits(), .lpBits, .Size)
            '-- Translate and decompose color
            Call OleTranslateColor(Color, 0, Color)
            R = (Color And &HFF&)
            G = (Color And &HFF00&) \ &H100
            B = (Color And &HFF0000) \ &H10000
            '-- Blend
            a2 = Amount
            a1 = &HFF - a2
            For i = 3 To .Size - 1 Step 4
                If (aBits(i)) Then 'is not transparent?
                    iIn = i - 1
                    aBits(iIn) = (a1 * aBits(iIn) + a2 * R) \ &HFF: iIn = iIn - 1
                    aBits(iIn) = (a1 * aBits(iIn) + a2 * G) \ &HFF: iIn = iIn - 1
                    aBits(iIn) = (a1 * aBits(iIn) + a2 * B) \ &HFF
                End If
            Next i
            '-- Unmap DIB bits
            Call pvUnmapDIBits(aBits())
        End If
    End With
End Sub

Public Sub ChangeGlobalAlpha(oDIB As cDIB32, ByVal GlobalAlpha As Byte)

  Dim uSA     As SAFEARRAY1D
  Dim aBits() As Byte
  Dim lAlpha  As Long
  Dim i       As Long
  
    With oDIB
        If (.hDIB) Then
            '-- Map DIB bits
            Call pvMapDIBits(uSA, aBits(), .lpBits, .Size)
            '-- Change gobal alpha
            lAlpha = GlobalAlpha
            For i = 3 To .Size - 1 Step 4
                If (aBits(i)) Then 'is not transparent?
                    aBits(i) = (aBits(i) * lAlpha) \ &HFF
                End If
            Next i
            '-- Unmap DIB bits
            Call pvUnmapDIBits(aBits())
        End If
    End With
End Sub

'//

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

