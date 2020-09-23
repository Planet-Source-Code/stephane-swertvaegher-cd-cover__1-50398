VERSION 5.00
Begin VB.Form CDC4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "CDC4.frx":0000
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Left            =   90
      TabIndex        =   0
      Top             =   3870
      Width           =   5820
   End
End
Attribute VB_Name = "CDC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
CDC4.Hide
End Sub

Private Sub Form_Load()
Me.Width = 400 * Screen.TwipsPerPixelX
Me.Height = 250 * Screen.TwipsPerPixelY
Label1 = "Contact me at:" & vbCrLf & "stephane.swertvaegher@pandora.be" & vbCrLf & vbCrLf & "Yes, I'm from Belgium, Europe !"
End Sub

Private Sub Label1_Click()
CDC4.Hide
End Sub
