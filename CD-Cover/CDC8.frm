VERSION 5.00
Begin VB.Form CDC8 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   3555
      Width           =   5640
      Begin VB.HScrollBar HScroll5 
         Height          =   240
         LargeChange     =   10
         Left            =   1350
         Max             =   500
         Min             =   -300
         TabIndex        =   12
         Top             =   2430
         Width           =   2745
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   240
         LargeChange     =   10
         Left            =   1350
         Max             =   500
         Min             =   -300
         TabIndex        =   11
         Top             =   2115
         Width           =   2745
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   240
         Left            =   1350
         Max             =   10
         Min             =   1
         TabIndex        =   10
         Top             =   1800
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
         Left            =   1350
         TabIndex        =   9
         Top             =   720
         Width           =   1500
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   240
         LargeChange     =   10
         Left            =   1350
         Max             =   300
         Min             =   10
         TabIndex        =   8
         Top             =   1350
         Value           =   10
         Width           =   2745
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   10
         Left            =   1350
         Max             =   300
         Min             =   10
         TabIndex        =   7
         Top             =   1035
         Value           =   10
         Width           =   2745
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mask color:"
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
         Left            =   2925
         TabIndex        =   24
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4185
         TabIndex        =   23
         Top             =   360
         Width           =   510
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
         Left            =   4140
         TabIndex        =   22
         Top             =   2430
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
         Left            =   495
         TabIndex        =   21
         Top             =   2430
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
         Left            =   4140
         TabIndex        =   20
         Top             =   2115
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
         Left            =   495
         TabIndex        =   19
         Top             =   2115
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
         Left            =   4140
         TabIndex        =   18
         Top             =   1800
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
         Left            =   495
         TabIndex        =   17
         Top             =   1800
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
         Left            =   4140
         TabIndex        =   16
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Left            =   4140
         TabIndex        =   15
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label12 
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
         Left            =   495
         TabIndex        =   14
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label Label13 
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
         Left            =   495
         TabIndex        =   13
         Top             =   1035
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Index           =   1
      Left            =   225
      TabIndex        =   25
      Top             =   3555
      Width           =   5640
      Begin VB.HScrollBar HScroll6 
         Height          =   240
         Left            =   1440
         Max             =   10
         Min             =   1
         TabIndex        =   27
         Top             =   585
         Value           =   1
         Width           =   2400
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 %"
         Height          =   285
         Left            =   3915
         TabIndex        =   28
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label14 
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
         Left            =   225
         TabIndex        =   26
         Top             =   585
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments for inlay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Index           =   2
      Left            =   225
      TabIndex        =   29
      Top             =   3555
      Width           =   5640
      Begin VB.HScrollBar HScroll7 
         Height          =   240
         Left            =   1440
         Max             =   10
         Min             =   1
         TabIndex        =   30
         Top             =   585
         Value           =   1
         Width           =   2400
      End
      Begin VB.Label Label17 
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
         Left            =   225
         TabIndex        =   32
         Top             =   585
         Width           =   1140
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 %"
         Height          =   285
         Left            =   3915
         TabIndex        =   31
         Top             =   585
         Width           =   780
      End
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   270
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   3
      Top             =   6570
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4770
      TabIndex        =   2
      Top             =   6570
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show me"
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   6570
      Width           =   1230
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   2940
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3690
      TabIndex        =   5
      Top             =   2700
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   3690
      Stretch         =   -1  'True
      Top             =   540
      Width           =   1950
   End
End
Attribute VB_Name = "CDC8"
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
If BgIdx = 0 Then
        .Pic3.Picture = Im
        GdiTransparentBlt .Pic3.hDC, .Shape1.Left, .Shape1.Top, .Shape1.Width, .Shape1.Height, Pic2.hDC, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight, Label2.BackColor
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll3.Value - 1)
End If
If BgIdx = 1 Then
    For xx = 0 To Int(.Pic1(PicIdx).Width / Pic2.Width)
    For yy = 0 To Int(.Pic1(PicIdx).Height / Pic2.Height)
    .Pic3.PaintPicture Pic2, xx * Pic2.Width, yy * Pic2.Height, Pic2.Width, Pic2.Height
    Next yy, xx
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll6.Value - 1)
End If
If BgIdx = 2 Then
    For xx = 0 To Int(.Pic1(PicIdx).Width / Pic2.Width)
    For yy = 0 To Int(.Pic1(PicIdx).Height / Pic2.Height)
    .Pic3.PaintPicture Pic2, xx * Pic2.Width, yy * Pic2.Height, Pic2.Width, Pic2.Height
    Next yy, xx
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll7.Value - 1)
        AlphaBlend .Pic1(PicIdx).hDC, .Pic3.Width - (SideW + 1), 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll7.Value - 1)
End If
        .Pic1(PicIdx).Refresh
End With
CDC1.Shape1.Visible = False
CDC8.Hide
End Sub

Private Sub Command2_Click() 'cancel
CDC1.Pic1(PicIdx) = Im
CDC1.Shape1.Visible = False
CDC8.Hide
End Sub

Private Sub Command3_Click() 'show me
With CDC1
        .Pic3.Picture = Nothing
        '.Pic4.Picture = Nothing
        .Pic1(PicIdx) = Im
        .Pic3.Width = .Pic1(PicIdx).Width
        .Pic3.Height = .Pic1(PicIdx).Height
If BgIdx = 0 Then
        .Pic3.Picture = Im
        GdiTransparentBlt .Pic3.hDC, .Shape1.Left, .Shape1.Top, .Shape1.Width, .Shape1.Height, Pic2.hDC, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight, Label2.BackColor
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll3.Value - 1)
End If
If BgIdx = 1 Then
    For xx = 0 To Int(.Pic1(PicIdx).Width / Pic2.Width)
    For yy = 0 To Int(.Pic1(PicIdx).Height / Pic2.Height)
    .Pic3.PaintPicture Pic2, xx * Pic2.Width, yy * Pic2.Height, Pic2.Width, Pic2.Height
    Next yy, xx
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll6.Value - 1)
End If
If BgIdx = 2 Then
    For xx = 0 To Int(.Pic1(PicIdx).Width / Pic2.Width)
    For yy = 0 To Int(.Pic1(PicIdx).Height / Pic2.Height)
    .Pic3.PaintPicture Pic2, xx * Pic2.Width, yy * Pic2.Height, Pic2.Width, Pic2.Height
    Next yy, xx
        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll7.Value - 1)
        AlphaBlend .Pic1(PicIdx).hDC, .Pic3.Width - (SideW + 1), 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll7.Value - 1)
End If
        .Pic1(PicIdx).Refresh
End With
End Sub

Private Sub File1_Click()
Pic2.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
Pic2.Refresh
Dimention2 Image1, Pic2, Pic2.ScaleWidth, Pic2.ScaleHeight
Label4 = Pic2.Width & " X " & Pic2.Height
CDC1.Shape1.Move (CDC1.Pic1(PicIdx).Width - Pic2.Width) / 2, (CDC1.Pic1(PicIdx).Height - Pic2.Height) / 2, Pic2.Width, Pic2.Height
HScroll4 = CDC1.Shape1.Left
HScroll5 = CDC1.Shape1.Top
HScroll1 = 100
HScroll2 = 100
End Sub

Private Sub Form_Activate()
On Error Resume Next
For xx = 0 To 4
Frame1(xx).Visible = False
Next xx
Set Im = CDC1.Pic1(PicIdx).Image
If BgIdx = 0 Then
Frame1(0).Visible = True
File1.Path = App.Path & "\CDC_Tubes"
File1.Pattern = "*.bmp; *.gif"
Set CDC1.Shape1.Container = CDC1.Pic1(PicIdx)
CDC1.Shape1.Visible = True
HScroll1 = 100
HScroll2 = 100
Check1.Value = 1
End If
If BgIdx = 1 Then
Frame1(1).Visible = True
File1.Path = App.Path & "\cdc-Textures"
File1.Pattern = "*.jpg; *.bmp; *.gif"
CDC1.Shape1.Visible = False
End If
If BgIdx = 2 Then
Frame1(2).Visible = True
File1.Path = App.Path & "\cdc-Textures"
File1.Pattern = "*.jpg; *.bmp; *.gif"
CDC1.Shape1.Visible = False
End If
File1.Selected(0) = True
End Sub

Private Sub Form_Load()
Me.Caption = CDCTitle
Me.Move 0, 450, 6180, 7410
T3D CDC8, Image1, 5, T3dRaiseInset
Label2.BackColor = 0
HScroll3 = 5
HScroll6 = 5
HScroll7 = 5
End Sub

Private Sub HScroll1_Change()
Label3 = Format(HScroll1 / 100, "0.00")
If Check1.Value = 1 Then HScroll2 = HScroll1
AdjustShape
End Sub

Private Sub HScroll2_Change()
Label5 = Format(HScroll2 / 100, "0.00")
If Check1.Value = 1 Then HScroll1 = HScroll2
AdjustShape
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

Private Sub HScroll6_Change()
Label15 = Format(HScroll6 * 10, "000") & " %"
End Sub

Private Sub HScroll7_Change()
Label16 = Format(HScroll7 * 10, "000") & " %"
End Sub

Private Sub Label2_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label2.BackColor
CDC2.CD1.ShowColor
Label2.BackColor = CDC2.CD1.Color

End Sub

Private Sub AdjustShape()
On Error Resume Next
With CDC1
NW = Pic2.Width * HScroll1 / 100
NH = Pic2.Height * HScroll2 / 100
.Shape1.Move (.Pic1(PicIdx).Width - NW) / 2, (.Pic1(PicIdx).Height - NH) / 2, NW, NH
HScroll4 = CDC1.Shape1.Left
HScroll5 = CDC1.Shape1.Top
End With
End Sub


