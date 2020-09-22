VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mAlphaBlt test (some effects)"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7515
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
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
   Begin VB.CheckBox chkClsBeforeRendering 
      Caption         =   "Cls before rendering"
      Height          =   240
      Left            =   5280
      TabIndex        =   11
      Top             =   4110
      Value           =   1  'Checked
      Width           =   1905
   End
   Begin VB.TextBox txtAlphaAmount 
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      Text            =   "128"
      Top             =   2745
      Width           =   540
   End
   Begin VB.CheckBox chkChangeGlobalAlpha 
      Caption         =   "Change global alpha"
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
      Left            =   5295
      TabIndex        =   8
      Top             =   2385
      Width           =   2040
   End
   Begin VB.CheckBox chkBlendWithColor 
      Caption         =   "Blend with color"
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
      Left            =   5295
      TabIndex        =   3
      Top             =   1140
      Width           =   1740
   End
   Begin VB.CheckBox chkConvertToGrey 
      Caption         =   "Convert to grey"
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
      Left            =   5295
      TabIndex        =   2
      Top             =   720
      Width           =   1740
   End
   Begin VB.TextBox txtBlendAmount 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Text            =   "64"
      Top             =   1905
      Width           =   540
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "&Cls"
      Height          =   435
      Left            =   5280
      TabIndex        =   12
      Top             =   4515
      Width           =   1005
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "&Render"
      Default         =   -1  'True
      Height          =   435
      Left            =   6375
      TabIndex        =   13
      Top             =   4515
      Width           =   1005
   End
   Begin VB.Label lblAlphaAmount 
      Caption         =   "Amount"
      Height          =   255
      Left            =   5580
      TabIndex        =   9
      Top             =   2790
      Width           =   795
   End
   Begin VB.Label lblBlendAmount 
      Caption         =   "Amount"
      Height          =   255
      Left            =   5580
      TabIndex        =   6
      Top             =   1950
      Width           =   795
   End
   Begin VB.Label lblBlendColorVal 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   6240
      TabIndex        =   5
      Top             =   1530
      Width           =   345
   End
   Begin VB.Label lblBlendColor 
      Caption         =   "Color"
      Height          =   240
      Left            =   5580
      TabIndex        =   4
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label lblBackColorVal 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   6240
      TabIndex        =   1
      Top             =   255
      Width           =   345
   End
   Begin VB.Label lblBackColor 
      Caption         =   "Back color"
      Height          =   240
      Left            =   5295
      TabIndex        =   0
      Top             =   255
      Width           =   1035
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

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

Private m_oDIB32       As cDIB32
Private m_oDIB32Effect As cDIB32



Private Sub Form_Load()
    
    If (App.LogMode <> 1) Then
        Call MsgBox("Absolutely recommended: compile first...")
    End If
    
    Set Me.Icon = Nothing
    
    '-- Load 32-bit bitmap
    Set m_oDIB32 = New cDIB32
    Call m_oDIB32.CreateFromBitmapFile(pvFixPath(App.Path) & "Test64x64x32bpp.bmp")
End Sub

Private Sub Form_Paint()
    Me.Line (0, 0)-(ScaleWidth, 0), vbButtonShadow
    Me.Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
    Me.Line (19, 19)-(340, 340), vbBlack, B
End Sub




Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub lblBackColorVal_Click()
    
  Dim lRet As Long
  
    lRet = mDialogColor.SelectColor(Me.hWnd, lblBackColorVal.BackColor, True)
    If (lRet <> -1) Then
        Me.FillColor = lRet
        Call Form_Paint
        lblBackColorVal.BackColor = lRet
    End If
End Sub

Private Sub lblBlendColorVal_Click()

  Dim lRet As Long
  Dim lClr As Long
    
    Call OleTranslateColor(lblBlendColorVal.BackColor, 0, lClr)
    lRet = mDialogColor.SelectColor(Me.hWnd, lClr, True)
    If (lRet <> -1) Then
        lblBlendColorVal.BackColor = lRet
    End If
End Sub

Private Sub cmdCls_Click()
    Call Form_Paint
End Sub

Private Sub cmdRender_Click()
    
  Dim i As Long
  
    With txtBlendAmount
        If (Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 255) Then
            Call MsgBox("Please, enter a valid 'Amount' number [0;255]")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
    End With
    With txtAlphaAmount
        If (Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 255) Then
            Call MsgBox("Please, enter a valid 'Amount' number [0;255]")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
    End With
    
    '-- Clone original bitmap
    Call m_oDIB32.CloneTo(m_oDIB32Effect)
    
    '-- Convert to grey (XP disabled effect)
    If (chkConvertToGrey) Then
        Call mDIB32Ext.ConvertToGrey(m_oDIB32Effect)
    End If
    '-- Blend with a given color (selection effect)
    If (chkBlendWithColor) Then
        Call mDIB32Ext.BlendWithColor(m_oDIB32Effect, lblBlendColorVal.BackColor, txtBlendAmount)
    End If
    '-- Change global-alpha (ghost effect)
    If (chkChangeGlobalAlpha) Then
        Call mDIB32Ext.ChangeGlobalAlpha(m_oDIB32Effect, txtAlphaAmount)
    End If
    
    '-- Render modified DIB...
    If (chkClsBeforeRendering) Then
        Call Form_Paint
    End If
    With m_oDIB32Effect
        Call AlphaBlend(Me.hDC, 20, 20, m_oDIB32.hDIB) 'reference
        For i = 1 To 24
            Call AlphaBlend(Me.hDC, 20 + (i Mod 5) * .Width, 20 + (i \ 5) * .Height, .hDIB)
        Next i
    End With
End Sub

Private Function pvFixPath(ByVal sPath As String) As String
    pvFixPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", vbNullString)
End Function
