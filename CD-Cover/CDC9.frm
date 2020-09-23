VERSION 5.00
Begin VB.Form CDC9 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Copy with shadow"
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
      TabIndex        =   4
      Top             =   4860
      Value           =   1  'Checked
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4275
      TabIndex        =   3
      Top             =   5355
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy !"
      Height          =   330
      Left            =   225
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   3615
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "CDC9.frx":0000
      Top             =   945
      Width           =   5685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter full text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   990
      TabIndex        =   0
      Top             =   315
      Width           =   3930
   End
End
Attribute VB_Name = "CDC9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EM_GETLINE = &HC4
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Command1_Click()
Temp = MsgBox("Copy the text to the " & Choose(PicIdx + 1, "frontside", "backside", "inlay", "label") & vbCrLf & "All existing text will be cleared." & vbCrLf & vbCrLf & "Are you sure about this ?", vbQuestion + vbYesNo, BLMTitle)
If Temp = vbNo Then Exit Sub
Select Case PicIdx
Case 0 ' copy to front
CopyTextLines CDC1.Lab0, CDC1.SLab0
Case 1 'copy to back
CopyTextLines CDC1.Lab1, CDC1.SLab1
End Select
CDC9.Hide
End Sub

Private Sub Command2_Click()
CDC9.Hide
End Sub

Private Sub Form_Load()
Me.Move 90, CDC1.Top, 6180, 6300
Me.Caption = CDCTitle
T3D CDC9, Label1, 5, T3dRaiseInset
T3D CDC9, Text1, 5, T3dRaiseInset
Text1 = ""
End Sub

Private Function GTBLine(objTB As TextBox, _
   ByVal LineNum As Integer) As String
    Dim lngRet As Long
    Dim lngLen As Long
    Dim lngFirstCharPos As Long
    Dim lngHwnd As Long
    Dim bytBuffer() As Byte
    Dim strAns As String

    If LineNum < 0 Then Exit Function
    
    If objTB.MultiLine = False Then
        GTBLine = objTB.Text
    Else
        lngHwnd = objTB.hWnd
        'first character position of the line
        lngFirstCharPos = SendMessage(lngHwnd, EM_LINEINDEX, _
              LineNum - 1, 0&)
        'length of line
        lngLen = SendMessage(lngHwnd, EM_LINELENGTH, _
           lngFirstCharPos, 0&)
        ReDim bytBuffer(lngLen) As Byte
        bytBuffer(0) = lngLen
        
        'text of line saved to bytBuffer
        lngRet = SendMessage(lngHwnd, EM_GETLINE, LineNum - 1, _
              bytBuffer(0))
        If lngRet Then strAns = Left$(StrConv(bytBuffer, _
            vbUnicode), lngLen)
        GTBLine = strAns
    End If
End Function

Private Sub CopyTextLines(Ob1 As Object, Ob2 As Object)
For xx = 0 To 49
    Ob1(xx).Visible = False
    Ob2(xx).Visible = False
If GTBLine(Text1, xx + 1) <> "" Then
    Ob1(xx) = GTBLine(Text1, xx + 1)
    Ob2(xx) = GTBLine(Text1, xx + 1)
    Ob1(xx).Visible = True
    If Check1.Value = 1 Then Ob2(xx).Visible = True
End If
Next xx
End Sub
