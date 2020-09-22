VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mAlphaBlt test"
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
   Begin VB.ComboBox cbScale 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6075
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2070
      Width           =   1080
   End
   Begin VB.TextBox txtIterations 
      Height          =   315
      Left            =   6375
      TabIndex        =   9
      Text            =   "1"
      Top             =   2940
      Width           =   660
   End
   Begin VB.OptionButton optTestFunction 
      Caption         =   "AlphaBlendStretch"
      Height          =   285
      Index           =   3
      Left            =   5385
      TabIndex        =   4
      Top             =   1695
      Width           =   1785
   End
   Begin VB.OptionButton optTestFunction 
      Caption         =   "AlphaBltStretch"
      Height          =   285
      Index           =   2
      Left            =   5385
      TabIndex        =   3
      Top             =   1320
      Width           =   1785
   End
   Begin VB.OptionButton optTestFunction 
      Caption         =   "AlphaBlend"
      Height          =   285
      Index           =   1
      Left            =   5385
      TabIndex        =   2
      Top             =   960
      Width           =   1785
   End
   Begin VB.OptionButton optTestFunction 
      Caption         =   "AlphaBlt"
      Height          =   285
      Index           =   0
      Left            =   5385
      TabIndex        =   1
      Top             =   585
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.CheckBox chkInterpolate 
      Caption         =   "Interpolate (bilinear)"
      Enabled         =   0   'False
      Height          =   210
      Left            =   5400
      TabIndex        =   7
      Top             =   2490
      Width           =   1890
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "&Render"
      Default         =   -1  'True
      Height          =   435
      Left            =   5835
      TabIndex        =   11
      Top             =   4515
      Width           =   1005
   End
   Begin VB.Label lblScale 
      Caption         =   "Scale to"
      Enabled         =   0   'False
      Height          =   240
      Left            =   5415
      TabIndex        =   5
      Top             =   2130
      Width           =   690
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
      TabIndex        =   8
      Top             =   2985
      Width           =   1020
   End
   Begin VB.Label lblTiming 
      Height          =   840
      Left            =   5415
      TabIndex        =   10
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Label lblTestFunction 
      Caption         =   "Test function:"
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
      Left            =   5385
      TabIndex        =   0
      Top             =   255
      Width           =   1740
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oDIB32 As cDIB32
Private m_oTile  As cTile
Private m_oT     As cTiming



Private Sub Form_Load()
    
    If (App.LogMode <> 1) Then
        Call MsgBox("Absolutely recommended: compile first...")
    End If
    
    Set Me.Icon = Nothing
    
    Set m_oDIB32 = New cDIB32
    Set m_oTile = New cTile
    Set m_oT = New cTiming
    
    '-- Load 32-bit bitmap
    Call m_oDIB32.CreateFromBitmapFile(pvFixPath(App.Path) & "Test64x64x32bpp.bmp")
   'Call m_oDIB32.CreateFromResourceBitmap([app. exe name], [res. ID])
        
    With cbScale
        Call .AddItem("1:2")
        Call .AddItem("1:1")
        Call .AddItem("2:1")
        Call .AddItem("3:1")
        Call .AddItem("4:1")
        Call .AddItem("5:1")
        Let .ListIndex = 5
    End With
End Sub

Private Sub Form_Paint()
    Me.Line (0, 0)-(ScaleWidth, 0), vbButtonShadow
    Me.Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
    Me.Line (19, 19)-(340, 340), vbBlack, B
End Sub



Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub mnuAbout_Click()
    Call MsgBox("Rendering alpha bitmaps" & vbCrLf & _
                "Carles P.V. - 2005" & vbCrLf & vbCrLf & _
                "Thanks to Peter Scale for 'Bilinear resizing' routine." & vbCrLf & _
                "Thanks to Ron van Tilburg for the 'integer maths' version" & vbCrLf & _
                "as well as for the 'Nearest neighbour resizing' routine.")
End Sub



Private Sub optTestFunction_Click(Index As Integer)
    chkInterpolate.Enabled = (Index >= 2)
    lblScale.Enabled = (Index >= 2)
    cbScale.Enabled = (Index >= 2)
End Sub

Private Sub cmdRender_Click()
    
  Dim SW As Long
  Dim SH As Long
  Dim it As Long
  Dim i  As Long
  Dim s  As String
    
    With txtIterations
        If (Not IsNumeric(.Text)) Then
            Call MsgBox("Please, enter a valid 'Iterations' number")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
    End With
    
    '-- Scale to
    Select Case True
      Case optTestFunction(2), optTestFunction(3)
        i = Abs(cbScale.ListIndex - 1) + 1
        With m_oDIB32
            If (cbScale.ListIndex < 1) Then
                SW = .Width \ i
                SH = .Height \ i
              Else
                SW = .Width * i
                SH = .Height * i
            End If
        End With
    End Select
    
    '-- Background
    Select Case True
      Case optTestFunction(2)
        Call m_oTile.CreatePatternFromSolidColor(vbCyan)
      Case optTestFunction(1), optTestFunction(3)
        Call m_oTile.CreatePatternFromStdPicture(LoadResPicture(101, vbResBitmap))
    End Select
    Call m_oTile.Tile(Me.hDC, 20, 20, 320, 320)
    
    '-- Render...
    Screen.MousePointer = vbArrowHourglass
    
    Call m_oT.Reset
    
    For it = 1 To txtIterations
        
        With m_oDIB32
        
            s = " " & .Width & "x" & .Height & "-bitmaps "
        
            Select Case True
              
              Case optTestFunction(0) 'AlphaBlt
                
                For i = 0 To 24
                    Call AlphaBlt(Me.hDC, 20 + (i Mod 5) * .Width, 20 + (i \ 5) * .Height, vbCyan, .hDIB)
                Next i
                s = it * 25 & s & "rendered in "
              
              Case optTestFunction(1) 'AlphaBlend
                
                For i = 0 To 24
                    Call AlphaBlend(Me.hDC, 20 + (i Mod 5) * .Width, 20 + (i \ 5) * .Height, .hDIB)
                Next i
                s = it * 25 & s & "rendered in "
              
              Case optTestFunction(2) 'AlphaBltStretch
                
                Call AlphaBltStretch(Me.hDC, 20, 20, SW, SH, vbCyan, chkInterpolate, .hDIB)
                s = it & s & "scaled and rendered in "
              
              Case optTestFunction(3) 'AlphaBlendStretch
                
                Call AlphaBlendStretch(Me.hDC, 20, 20, SW, SH, chkInterpolate, .hDIB)
                s = it & s & "scaled and rendered in "
            End Select
        End With
    Next it
    
    lblTiming.Caption = s & Format$(m_oT.Elapsed / 1000, "0.000 sec.")
    
    Screen.MousePointer = vbDefault
End Sub

Private Function pvFixPath(ByVal sPath As String) As String
    pvFixPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", vbNullString)
End Function
