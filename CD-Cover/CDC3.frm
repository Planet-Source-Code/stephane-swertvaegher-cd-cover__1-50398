VERSION 5.00
Begin VB.Form CDC3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Show me"
      Height          =   375
      Left            =   135
      TabIndex        =   19
      Top             =   3015
      Width           =   1230
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   240
      LargeChange     =   10
      Left            =   1035
      Max             =   500
      Min             =   -300
      TabIndex        =   17
      Top             =   2520
      Width           =   2745
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   240
      LargeChange     =   10
      Left            =   1035
      Max             =   500
      Min             =   -300
      TabIndex        =   14
      Top             =   2205
      Width           =   2745
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   240
      Left            =   1035
      Max             =   10
      Min             =   1
      TabIndex        =   11
      Top             =   1890
      Value           =   1
      Width           =   2745
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Lock scales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1035
      TabIndex        =   9
      Top             =   720
      Width           =   1500
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   240
      LargeChange     =   10
      Left            =   1035
      Max             =   300
      Min             =   10
      TabIndex        =   6
      Top             =   1350
      Value           =   10
      Width           =   2745
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   1035
      Max             =   300
      Min             =   10
      TabIndex        =   5
      Top             =   1035
      Value           =   10
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3330
      TabIndex        =   1
      Top             =   3015
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1980
      TabIndex        =   0
      Top             =   3015
      Width           =   1230
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3825
      TabIndex        =   18
      Top             =   2520
      Width           =   645
   End
   Begin VB.Label Label10 
      Caption         =   "Offset Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   16
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3825
      TabIndex        =   15
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label Label8 
      Caption         =   "Offset X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   13
      Top             =   2205
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3825
      TabIndex        =   12
      Top             =   1890
      Width           =   645
   End
   Begin VB.Label Label6 
      Caption         =   "Alphabl:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   1890
      Width           =   780
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3825
      TabIndex        =   8
      Top             =   1350
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3825
      TabIndex        =   7
      Top             =   1035
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Scale Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Scale X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copy scaled picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   765
      TabIndex        =   2
      Top             =   225
      Width           =   3120
   End
End
Attribute VB_Name = "CDC3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'ok
CDC1.Pic1(PicIdx) = Im
SaveRedo
AlphaBlend CDC1.Pic1(PicIdx).hDC, CDC1.Shape1.Left, CDC1.Shape1.Top, CDC1.Shape1.Width, CDC1.Shape1.Height, CDC1.Pic2.hDC, 0, 0, CDC1.Pic2.Width, CDC1.Pic2.Height, Alpha(HScroll3.Value - 1)
CDC1.Pic1(PicIdx).Refresh
CDC1.Shape1.Visible = False
CDC3.Hide
End Sub

Private Sub Command2_Click() 'cancel
CDC1.Pic1(PicIdx) = Im
CDC1.Shape1.Visible = False
CDC3.Hide
End Sub

Private Sub Command3_Click() 'show me
CDC1.Pic1(PicIdx) = Im
AlphaBlend CDC1.Pic1(PicIdx).hDC, CDC1.Shape1.Left, CDC1.Shape1.Top, CDC1.Shape1.Width, CDC1.Shape1.Height, CDC1.Pic2.hDC, 0, 0, CDC1.Pic2.Width, CDC1.Pic2.Height, Alpha(HScroll3.Value - 1)
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub Form_Activate()
HScroll1 = 100
HScroll2 = 100
Check1.Value = 1
Set CDC1.Shape1.Container = CDC1.Pic1(PicIdx)
CDC1.Shape1.Move (CDC1.Pic1(PicIdx).Width - CDC1.Pic2.Width) / 2, (CDC1.Pic1(PicIdx).Height - CDC1.Pic2.Height) / 2, CDC1.Pic2.Width, CDC1.Pic2.Height
CDC1.Shape1.Visible = True
Set Im = CDC1.Pic1(PicIdx).Image
End Sub

Private Sub Form_Load()
Me.Caption = CDCTitle
Me.Move 0, CDC1.Top, 4770, 3900
T3D CDC3, CDC3.Label1, 5, T3dInsetRaise
HScroll3 = 5
End Sub

Private Sub HScroll1_Change()
Label4 = Format(HScroll1 / 100, "0.00")
If Check1.Value = 1 Then HScroll2 = HScroll1
AdjustShape
End Sub

Private Sub HScroll2_Change()
Label5 = Format(HScroll2 / 100, "0.00")
If Check1.Value = 1 Then HScroll1 = HScroll2
AdjustShape
End Sub

Private Sub AdjustShape()
On Error Resume Next
With CDC1
NW = .Pic2.Width * HScroll1 / 100
NH = .Pic2.Height * HScroll2 / 100
.Shape1.Move (.Pic1(PicIdx).Width - NW) / 2, (.Pic1(PicIdx).Height - NH) / 2, NW, NH
HScroll4 = CDC1.Shape1.Left
HScroll5 = CDC1.Shape1.Top
End With
End Sub

Private Sub HScroll3_Change()
Label7 = Format(HScroll3 * 10, "000") & " %"
End Sub

Private Sub HScroll4_Change()
Label9 = Format(HScroll4, "000")
CDC1.Shape1.Left = HScroll4
End Sub

Private Sub HScroll5_Change()
Label11 = Format(HScroll5, "000")
CDC1.Shape1.Top = HScroll5
End Sub
