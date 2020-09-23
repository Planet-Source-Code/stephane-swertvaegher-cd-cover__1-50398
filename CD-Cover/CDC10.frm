VERSION 5.00
Begin VB.Form CDC10 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   2715
   ClientTop       =   2430
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4770
      TabIndex        =   9
      Top             =   7785
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select font"
      Height          =   420
      Left            =   3375
      TabIndex        =   8
      Top             =   7785
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   1845
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5325
      LargeChange     =   25
      Left            =   5715
      TabIndex        =   1
      Top             =   540
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   45
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   0
      Top             =   540
      Width           =   5685
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   0
         ScaleHeight     =   262
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   328
         TabIndex        =   3
         Top             =   0
         Width           =   4920
         Begin VB.Label Label1 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3390
         End
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   5805
      Width           =   6000
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   90
      TabIndex        =   6
      Top             =   6210
      Width           =   5955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Number of fonts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   945
      TabIndex        =   5
      Top             =   135
      Width           =   4065
   End
End
Attribute VB_Name = "CDC10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lbc As Long, LabelIdx%

Private Sub Command1_Click()
'select & hide
If Showfont = 0 Then
CDC2.Combo1.Text = Label1(LabelIdx)
    If PicIdx = 0 Then CDC1.Lab0(q).Font = CDC2.Combo1.Text
    If PicIdx = 1 Then CDC1.Lab1(q).Font = CDC2.Combo1.Text
    If PicIdx = 0 Then CDC1.SLab0(q).Font = CDC2.Combo1.Text
    If PicIdx = 1 Then CDC1.SLab1(q).Font = CDC2.Combo1.Text
CDC1.Pic1(PicIdx).Refresh
End If
CDC10.Hide
End Sub

Private Sub Command2_Click()
CDC10.Hide
End Sub

Private Sub Form_Activate()
For xx = 0 To Label1.UBound
Label1(xx).BackColor = lbc
Next xx
If Showfont = 0 Then 'from cdc2
    Label3.Font = CDC2.Combo1.Text
    Label4 = CDC2.Combo1.Text
        For xx = 0 To Label1.UBound
            If Label1(xx) = CDC2.Combo1.Text Then
                Label1(xx).BackColor = vbBlue
                LabelIdx = xx
                Exit For
            End If
        Next xx
                If (xx * -1) < VScroll1.Max Then
                VScroll1 = VScroll1.Max
                Else
                VScroll1.Value = (xx * -1)
                End If
End If
End Sub

Private Sub Form_Load()
Me.Move 0, 0
Me.Caption = CDCTitle
Label1(0).Move 0, 0, Picture1.ScaleWidth - 2, 19
lbc = Label1(0).BackColor
Picture1.Height = 18 * Label1(0).Height
VScroll1.Height = Picture1.Height
T3D CDC10, Label2, 5, T3dRaiseInset
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Command2_Click
End Sub

Private Sub Label1_Click(Index As Integer)
For xx = 0 To Label1.UBound
Label1(xx).BackColor = lbc
Next xx
Label3.Font = List1.List(Index)
Label4 = List1.List(Index)
Label1(Index).BackColor = vbBlue
LabelIdx = Index
End Sub

Private Sub VScroll1_Change()
Picture2.Top = VScroll1.Value * Label1(0).Height
End Sub
