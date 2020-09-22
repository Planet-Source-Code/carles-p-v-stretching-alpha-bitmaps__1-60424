VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GDI+ test"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7515
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optBackground 
      Caption         =   "Pattern"
      Height          =   255
      Index           =   1
      Left            =   5385
      TabIndex        =   2
      Top             =   870
      Width           =   1545
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "Solid"
      Height          =   255
      Index           =   0
      Left            =   5385
      TabIndex        =   1
      Top             =   555
      Value           =   -1  'True
      Width           =   705
   End
   Begin VB.TextBox txtIterations 
      Height          =   315
      Left            =   6375
      TabIndex        =   5
      Text            =   "1"
      Top             =   1875
      Width           =   660
   End
   Begin VB.CheckBox chkInterpolate 
      Caption         =   "Interpolate (bilinear)"
      Height          =   210
      Left            =   5400
      TabIndex        =   3
      Top             =   1365
      Width           =   1890
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "&Render"
      Default         =   -1  'True
      Height          =   435
      Left            =   5835
      TabIndex        =   7
      Top             =   4515
      Width           =   1005
   End
   Begin VB.Label lblSolidColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   6105
      TabIndex        =   8
      Top             =   555
      Width           =   345
   End
   Begin VB.Label lblBackground 
      Caption         =   "Background:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Top             =   255
      Width           =   1995
   End
   Begin VB.Label lblIterations 
      Caption         =   "Iterations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5415
      TabIndex        =   4
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label lblTiming 
      Height          =   840
      Left            =   5415
      TabIndex        =   6
      Top             =   2490
      Width           =   1905
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_hGDIpToken As Long

Private m_oDIB32     As cDIB32
Private m_oTile      As cTile
Private m_oT         As cTiming



Private Sub Form_Load()
  
  Dim GDIpInput As GdiplusStartupInput
    
    Set Me.Icon = Nothing
    
    '-- Initialize the GDI+ library
    GDIpInput.GdiplusVersion = 1
    If (mGDIp.GdiplusStartup(m_hGDIpToken, GDIpInput) <> [OK]) Then
        GoTo errH
    End If

    
    Set m_oDIB32 = New cDIB32
    Set m_oTile = New cTile
    Set m_oT = New cTiming
    
    '-- Load 32-bit bitmap
    Call m_oDIB32.CreateFromBitmapFile(pvFixPath(App.Path) & "Test64x64x32bpp.bmp")
   'Call m_oDIB32.CreateFromResourceBitmap([app. exe name], [res. ID])
    Exit Sub
    
errH:
    Call MsgBox("Error loading GDI+!", vbCritical)
    Call Unload(Me)
    On Error GoTo 0
End Sub

Private Sub Form_Paint()
    Me.Line (0, 0)-(ScaleWidth, 0), vbButtonShadow
    Me.Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
    Me.Line (19, 19)-(340, 340), vbBlack, B
End Sub



Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (m_hGDIpToken <> 0) Then
        Call mGDIp.GdiplusShutdown(m_hGDIpToken)
    End If
End Sub



Private Sub lblSolidColor_Click()
    
  Dim lRet As Long
  
    lRet = mDialogColor.SelectColor(Me.hWnd, lblSolidColor.BackColor, True)
    If (lRet <> -1) Then
        lblSolidColor.BackColor = lRet
    End If
End Sub

Private Sub cmdRender_Click()
    
  Dim it As Long
  Dim i  As Long
    
    With txtIterations
        If (Not IsNumeric(.Text)) Then
            Call MsgBox("Please, enter a valid 'Iterations' number")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbArrowHourglass
    
    Select Case True
      Case optBackground(0)
        Call m_oTile.CreatePatternFromSolidColor(lblSolidColor.BackColor)
      Case optBackground(1)
        Call m_oTile.CreatePatternFromStdPicture(LoadResPicture(101, vbResBitmap))
    End Select
    Call m_oTile.Tile(Me.hDC, 20, 20, 320, 320)
    
    Call m_oT.Reset
    
    For it = 1 To txtIterations
        Call mGDIp.StretchDIB32(m_oDIB32, Me.hDC, 20, 20, 320, 320, , , , , , , chkInterpolate)
    Next it
    
    lblTiming.Caption = "GDI+ rendering in " & Format$(m_oT.Elapsed / 1000, "0.000 sec.")
    
    Screen.MousePointer = vbDefault
End Sub

Private Function pvFixPath(ByVal sPath As String) As String
    pvFixPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", vbNullString)
End Function
