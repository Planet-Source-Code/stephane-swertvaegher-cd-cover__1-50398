VERSION 5.00
Begin VB.Form CDC7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show me"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Borders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5865
      Begin VB.HScrollBar HScroll4 
         Height          =   240
         LargeChange     =   10
         Left            =   1620
         Max             =   500
         TabIndex        =   22
         Top             =   2520
         Value           =   1
         Width           =   2805
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   240
         LargeChange     =   10
         Left            =   1620
         Max             =   500
         TabIndex        =   20
         Top             =   2205
         Value           =   1
         Width           =   2805
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   240
         LargeChange     =   10
         Left            =   1620
         Max             =   250
         Min             =   1
         TabIndex        =   18
         Top             =   1890
         Value           =   1
         Width           =   2805
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Equal border"
         Height          =   240
         Index           =   0
         Left            =   1665
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gradient border #1"
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   7
         Top             =   540
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gradient border #2"
         Height          =   240
         Index           =   2
         Left            =   1665
         TabIndex        =   6
         Top             =   810
         Width           =   1785
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   1620
         Max             =   10
         Min             =   1
         TabIndex        =   5
         Top             =   1575
         Value           =   1
         Width           =   2805
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         Height          =   285
         Left            =   135
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   4
         Top             =   1170
         Width           =   2220
      End
      Begin VB.Label Label10 
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
         Left            =   4545
         TabIndex        =   23
         Top             =   2520
         Width           =   645
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
         Left            =   4545
         TabIndex        =   21
         Top             =   2205
         Width           =   645
      End
      Begin VB.Label Label8 
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
         Left            =   4545
         TabIndex        =   19
         Top             =   1890
         Width           =   645
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   2475
         Width           =   1275
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borderwidth:"
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
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1845
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color 1"
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
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color 2"
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
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   12
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   11
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alphablend:"
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
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 %"
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
         Left            =   4545
         TabIndex        =   9
         Top             =   1575
         Width           =   645
      End
   End
End
Attribute VB_Name = "CDC7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'ok
With CDC1
.Pic3.Picture = Nothing
'.Pic4.Picture = Nothing
.Pic1(PicIdx) = Im
SaveRedo
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
If Option1(0).Value = True Then Border0
If Option1(1).Value = True Then Border1 Label3.BackColor, Label4.BackColor, HScroll2
If Option1(2).Value = True Then Border2 Label3.BackColor, Label4.BackColor, HScroll2
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll1.Value - 1)
.Pic1(PicIdx).Refresh
CDC7.Hide
End With
End Sub

Private Sub Command2_Click() 'cancel
CDC1.Pic1(PicIdx) = Im
CDC7.Hide
End Sub

Private Sub Command3_Click() 'show me
With CDC1
.Pic3.Picture = Nothing
'.Pic4.Picture = Nothing
.Pic1(PicIdx) = Im
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
If Option1(0).Value = True Then Border0
If Option1(1).Value = True Then Border1 Label3.BackColor, Label4.BackColor, HScroll2
If Option1(2).Value = True Then Border2 Label3.BackColor, Label4.BackColor, HScroll2
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll1.Value - 1)
.Pic1(PicIdx).Refresh
End With
End Sub

Private Sub Form_Activate()
Set Im = CDC1.Pic1(PicIdx).Image
End Sub

Private Sub Form_Load()
Me.Caption = CDCTitle
Me.Move 0, 450, 6180, 3945
HScroll1 = 5
HScroll2 = 30
HScroll3 = 0
HScroll4 = 0
    Pic1.BackColor = Label3.BackColor
End Sub

Private Sub HScroll1_Change()
Label6 = Format(HScroll1 * 10, "000") & " %"
End Sub

Private Sub HScroll2_Change()
Label8 = Format(HScroll2, "000")
End Sub

Private Sub HScroll3_Change()
Label9 = Format(HScroll3, "000")
End Sub

Private Sub HScroll4_Change()
Label10 = Format(HScroll4, "000")
End Sub

Private Sub Border0()
With CDC1
For xx = 0 To HScroll2.Value - 1
.Pic3.Line (HScroll3 + xx, HScroll4 + xx)-(.Pic3.Width - 1 - xx - HScroll3, .Pic3.Height - 1 - xx - HScroll4), Label3.BackColor, B
Next xx
End With
End Sub

Private Sub Border1(Col1&, Col2&, Dist%) 'add grad border 1
On Error Resume Next
With CDC1
Dim R1%, G1%, B1%, R2%, G2%, B2%, Sr, Sg, Sb
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
Sr = (R2 - R1) / Dist
Sg = (G2 - G1) / Dist
Sb = (B2 - B1) / Dist
For xx = 0 To Dist
.Pic3.Line (HScroll3 + xx, HScroll4 + xx)-(.Pic3.Width - 1 - xx - HScroll3, .Pic3.Height - 1 - xx - HScroll4), RGB(R1, G1, B1), B
R1 = R1 + Sr
G1 = G1 + Sg
B1 = B1 + Sb
Next xx
End With
End Sub

Private Sub Border2(Col1&, Col2&, Dist%) 'add grad border 1
On Error Resume Next
With CDC1
Dim R1%, G1%, B1%, R2%, G2%, B2%, Sr, Sg, Sb
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
Sr = (R2 - R1) / (Dist / 2)
Sg = (G2 - G1) / (Dist / 2)
Sb = (B2 - B1) / (Dist / 2)
For xx = 0 To Dist / 2
.Pic3.Line (HScroll3 + xx, HScroll4 + xx)-(.Pic3.Width - 1 - xx - HScroll3, .Pic3.Height - 1 - xx - HScroll4), RGB(R1, G1, B1), B
.Pic3.Line (HScroll3 + Dist - xx, HScroll4 + Dist - xx)-(.Pic3.Width - HScroll3 - Dist + xx, .Pic3.Height - HScroll4 - Dist + xx), RGB(R1, G1, B1), B
R1 = R1 + Sr
G1 = G1 + Sg
B1 = B1 + Sb
If R1 > 255 Then R1 = 255
If G1 > 255 Then G1 = 255
If B1 > 255 Then B1 = 255
If R1 < 0 Then R1 = 0
If B1 < 0 Then B1 = 0
If G1 < 0 Then G1 = 0
Next xx
End With
End Sub

Private Sub Label3_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label3.BackColor
CDC2.CD1.ShowColor
Label3.BackColor = CDC2.CD1.Color
    If Option1(0).Value = True Then
    Pic1.BackColor = Label3.BackColor
    End If
    If Option1(1).Value = True Then
    Grad Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
    If Option1(2).Value = True Then
    Grad2 Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
End Sub

Private Sub Label4_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label4.BackColor
CDC2.CD1.ShowColor
Label4.BackColor = CDC2.CD1.Color
    If Option1(0).Value = True Then
    Pic1.BackColor = Label3.BackColor
    End If
    If Option1(1).Value = True Then
    Grad Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
    If Option1(2).Value = True Then
    Grad2 Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
    Pic1.BackColor = Label3.BackColor
    End If
    If Option1(1).Value = True Then
    Grad Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
    If Option1(2).Value = True Then
    Grad2 Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
End Sub
