VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CDC6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   638
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   933
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Save printpositions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   12
      Top             =   675
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9810
      TabIndex        =   9
      Text            =   "000"
      Top             =   1080
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9810
      TabIndex        =   8
      Text            =   "000"
      Top             =   720
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   225
      TabIndex        =   5
      Top             =   1215
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print !"
      Height          =   420
      Left            =   225
      TabIndex        =   4
      Top             =   675
      Width           =   1635
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   3
      Left            =   5220
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   3
      Top             =   1980
      Width           =   1455
      Begin VB.Image Image2 
         Height          =   6855
         Left            =   315
         Picture         =   "CDC6.frx":0000
         Top             =   180
         Width           =   6855
      End
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   3600
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   2
      Top             =   1980
      Width           =   1455
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   1980
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   1980
      Width           =   1455
   End
   Begin VB.PictureBox Image1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   0
      Left            =   360
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   1935
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   225
      TabIndex        =   6
      Top             =   90
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kFront"
            Object.ToolTipText     =   "Text of the frontside"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kBack"
            Object.ToolTipText     =   "Text of the backside"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kInlay"
            Object.ToolTipText     =   "Text of the inlay"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kLabel"
            Object.ToolTipText     =   "Text of the round label"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Y Position:"
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
      Left            =   8460
      TabIndex        =   11
      Top             =   1125
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "X Position:"
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
      Left            =   8460
      TabIndex        =   10
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Height          =   330
      Left            =   2925
      TabIndex        =   7
      Top             =   1440
      Width           =   3480
   End
End
Attribute VB_Name = "CDC6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrtIdx%, BCCol As Long

Private Sub Command1_Click()
Select Case PrtIdx
Case 0
Temp = MsgBox("This will print the front- and backside." & vbCrLf & vbCrLf & "Continue ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
PrintFront
Case 1
Temp = MsgBox("This will print the front- and backside." & vbCrLf & vbCrLf & "Continue ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
PrintFront
Case 2
Temp = MsgBox("This will print the inlay." & vbCrLf & vbCrLf & "Continue ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
PrintInlay
Case 3
Temp = MsgBox("This will print the label." & vbCrLf & vbCrLf & "Continue ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
PrintLabel
End Select
End Sub

Private Sub Command2_Click()
CDC6.Hide
End Sub

Private Sub Command3_Click()
On Error GoTo C3
Temp = MsgBox("This will save the print-positions" & vbCrLf & "of the booklet, inlay and label." & vbCrLf & vbCrLf & "Continue ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
ff = FreeFile
Open App.Path & "\cdc_Pref\prtpos.txt" For Output As #ff
For xx = 0 To 2
Print #ff, PrtposX(xx)
Print #ff, PrtposY(xx)
Next xx
Close #ff
Exit Sub
C3:
MsgBox Err.Description, vbCritical + vbOKOnly, CDCTitle
Close #ff
End Sub

Private Sub Form_Activate()
Image1(0).Picture = Nothing
Image1(1).Picture = Nothing
'Image1(2).Picture = Nothing
'Image1(3).Picture = Nothing
Image1(0).Picture = CDC1.Pic1(0).Image
Image1(1).Picture = CDC1.Pic1(1).Image
'Image1(2).Picture = CDC1.Pic1(2).Image
'Image1(3).Picture = CDC1.Pic1(3).Image
If PicIdx = 0 Then SetImages True, True, False, False
If PicIdx = 1 Then SetImages True, True, False, False
'If PicIdx = 2 Then SetImages False, False, True, False
'If PicIdx = 3 Then SetImages False, False, False, True
PrtIdx = PicIdx
With CDC1
For xx = 0 To 49
    If .SLab0(xx).Visible = True Then
    Image1(0).Font = .SLab0(xx).Font
    Image1(0).FontSize = .SLab0(xx).FontSize
    Image1(0).ForeColor = .SLab0(xx).ForeColor
    Image1(0).CurrentX = .SLab0(xx).Left
    Image1(0).CurrentY = .SLab0(xx).Top
    Image1(0).Print .SLab0(xx)
    End If
    If .SLab1(xx).Visible = True Then
    Image1(1).Font = .SLab1(xx).Font
    Image1(1).FontSize = .SLab1(xx).FontSize
    Image1(1).ForeColor = .SLab1(xx).ForeColor
    Image1(1).CurrentX = .SLab1(xx).Left
    Image1(1).CurrentY = .SLab1(xx).Top
    Image1(1).Print .SLab1(xx)
    End If
    If .SLab2(xx).Visible = True Then
    Image1(2).Font = .SLab2(xx).Font
    Image1(2).FontSize = .SLab2(xx).FontSize
    Image1(2).ForeColor = .SLab2(xx).ForeColor
    Image1(2).CurrentX = .SLab2(xx).Left
    Image1(2).CurrentY = .SLab2(xx).Top
    Image1(2).Print .SLab2(xx)
    End If
    If .SLab3(xx).Visible = True Then
    Image1(3).Font = .SLab3(xx).Font
    Image1(3).FontSize = .SLab3(xx).FontSize
    Image1(3).ForeColor = .SLab3(xx).ForeColor
    Image1(3).CurrentX = .SLab3(xx).Left
    Image1(3).CurrentY = .SLab3(xx).Top
    Image1(3).Print .SLab3(xx)
    End If
Next xx
For xx = 0 To 49
    If .Lab0(xx).Visible = True Then
    Image1(0).Font = .Lab0(xx).Font
    Image1(0).FontSize = .Lab0(xx).FontSize
    Image1(0).ForeColor = .Lab0(xx).ForeColor
    Image1(0).CurrentX = .Lab0(xx).Left
    Image1(0).CurrentY = .Lab0(xx).Top
    Image1(0).Print .Lab0(xx)
    End If
    If .Lab1(xx).Visible = True Then
    Image1(1).Font = .Lab1(xx).Font
    Image1(1).FontSize = .Lab1(xx).FontSize
    Image1(1).ForeColor = .Lab1(xx).ForeColor
    Image1(1).CurrentX = .Lab1(xx).Left
    Image1(1).CurrentY = .Lab1(xx).Top
    Image1(1).Print .Lab1(xx)
    End If
    If .Lab2(xx).Visible = True Then
    Image1(2).Font = .Lab2(xx).Font
    Image1(2).FontSize = .Lab2(xx).FontSize
    Image1(2).ForeColor = .Lab2(xx).ForeColor
    Image1(2).CurrentX = .Lab2(xx).Left
    Image1(2).CurrentY = .Lab2(xx).Top
    Image1(2).Print .Lab2(xx)
    End If
    If .Lab3(xx).Visible = True Then
    Image1(3).Font = .Lab3(xx).Font
    Image1(3).FontSize = .Lab3(xx).FontSize
    Image1(3).ForeColor = .Lab3(xx).ForeColor
    Image1(3).CurrentX = .Lab3(xx).Left
    Image1(3).CurrentY = .Lab3(xx).Top
    Image1(3).Print .Lab3(xx)
    End If
Next xx
Toolbar2.Buttons(1).Value = tbrUnpressed
Toolbar2.Buttons(2).Value = tbrUnpressed
Toolbar2.Buttons(3).Value = tbrUnpressed
Toolbar2.Buttons(4).Value = tbrUnpressed
Toolbar2.Buttons(PicIdx + 1).Value = tbrPressed
Label1 = "Print " & Choose(PicIdx + 1, "front- and backside", "Front- and backside", "Inlay", "label")
End With
If PrtIdx = 0 Or PrtIdx = 1 Then
Text1.Text = PrtposX(0)
Text2.Text = PrtposY(0)
End If
If PrtIdx = 2 Then
Text1.Text = PrtposX(1)
Text2.Text = PrtposY(1)
End If
If PrtIdx = 3 Then
Text1.Text = PrtposX(2)
Text2.Text = PrtposY(2)
End If
If PrtIdx = 3 Then
Image2.Visible = True
Else
Image2.Visible = False
End If
End Sub

Private Sub Form_Load()
On Error GoTo FLoad2
Me.Caption = CDCTitle
Image1(0).Move ((CDC6.ScaleWidth - 457 - 457) / 2) + 458, 140, 457, 457
Image1(1).Move (CDC6.ScaleWidth - 457 - 457) / 2, 140, 457, 457
'Image1(2).Move (CDC6.ScaleWidth - PicW(2)) / 2, 145, PicW(2), PicH(2)
'Image1(3).Move (CDC6.ScaleWidth - 457) / 2, 140, 457, 457
Image2.Move 0, 0, 457, 457
Toolbar2.ImageList = CDC1.ImageList1
Toolbar2.Buttons(1).Image = 2
Toolbar2.Buttons(2).Image = 3
Toolbar2.Buttons(3).Image = 4
Toolbar2.Buttons(4).Image = 5
Label1.Move (CDC6.ScaleWidth - Label1.Width) / 2, 10, 232
Command3.Move (CDC6.ScaleWidth - Command3.Width) / 2, 60
ff = FreeFile
Open App.Path & "\cdc_Pref\prtpos.txt" For Input As #ff
For xx = 0 To 2
Input #ff, PrtposX(xx)
Input #ff, PrtposY(xx)
Next xx
Close #ff
Exit Sub
FLoad2:
MsgBox Err.Description, vbCritical + vbOKOnly, CDCTitle
Close #ff
End Sub

Private Sub SetImages(I0 As Boolean, I1 As Boolean, I2 As Boolean, I3 As Boolean)
Image1(0).Visible = I0
Image1(1).Visible = I1
Image1(2).Visible = I2
Image1(3).Visible = I3
End Sub

Private Sub Text1_Change()
If PrtIdx = 0 Or PrtIdx = 1 Then
PrtposX(0) = Text1.Text
End If
If PrtIdx = 2 Then
PrtposX(1) = Text1.Text
End If
If PrtIdx = 3 Then
PrtposX(2) = Text1.Text
End If
End Sub

Private Sub Text2_Change()
If PrtIdx = 0 Or PrtIdx = 1 Then
PrtposY(0) = Text2.Text
End If
If PrtIdx = 2 Then
PrtposY(1) = Text2.Text
End If
If PrtIdx = 3 Then
PrtposY(2) = Text2.Text
End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Image2.Visible = False
If Toolbar2.Buttons(1).Value = tbrPressed Then
SetImages True, True, False, False
    Label1 = "Print front- and backside"
    PrtIdx = 0
End If
If Toolbar2.Buttons(2).Value = tbrPressed Then
SetImages True, True, False, False
    Label1 = "Print front- and backside"
    PrtIdx = 1
End If
If Toolbar2.Buttons(3).Value = tbrPressed Then
SetImages False, False, True, False
    Label1 = "Print inlay"
    PrtIdx = 2
End If
If Toolbar2.Buttons(4).Value = tbrPressed Then
SetImages False, False, False, True
    Label1 = "Print label"
    PrtIdx = 3
End If
If PrtIdx = 0 Or PrtIdx = 1 Then
Text1.Text = PrtposX(0)
Text2.Text = PrtposY(0)
End If
If PrtIdx = 2 Then
Text1.Text = PrtposX(1)
Text2.Text = PrtposY(1)
End If
If PrtIdx = 3 Then
Text1.Text = PrtposX(2)
Text2.Text = PrtposY(2)
Image2.Visible = True
End If
End Sub
