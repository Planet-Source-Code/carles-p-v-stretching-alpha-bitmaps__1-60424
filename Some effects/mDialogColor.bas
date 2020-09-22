Attribute VB_Name = "mDialogColor"
'================================================
' Module:        mDialogColor.bas
' Author:        -
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type uCHOOSECOLOR
    lStructSize    As Long
    hWndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100
Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (Color As uCHOOSECOLOR) As Long

'//

Private m_nCustomColors(15) As Long
Private m_bInitialized      As Boolean



'========================================================================================
' Methods
'========================================================================================

Public Function SelectColor(ByVal hWndOwner As Long, ByVal DefaultColor As Long, Optional ByVal ShowExtendedDialog As Boolean = False) As Long
 
  Dim uCC  As uCHOOSECOLOR
  Dim lRet As Long
  Dim nIdx As Integer
     
    '-- Initialize custom colors
    If (m_bInitialized = False) Then
        m_bInitialized = True
        For nIdx = 0 To 15
            m_nCustomColors(nIdx) = QBColor(nIdx)
        Next nIdx
    End If
    
    '-- Prepare struct.
    With uCC
        .lStructSize = Len(uCC)
        .hWndOwner = hWndOwner
        .rgbResult = DefaultColor
        .lpCustColors = VarPtr(m_nCustomColors(0))
        .Flags = IIf(ShowExtendedDialog, CC_EXTENDED, CC_NORMAL)
    End With
        
    '-- Show Color dialog
    lRet = ChooseColor(uCC)
     
    '-- Get color / Cancel
    If (lRet) Then
        SelectColor = uCC.rgbResult
      Else
        SelectColor = -1
    End If
End Function
