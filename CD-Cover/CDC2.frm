VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CDC2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   225
      TabIndex        =   29
      Top             =   135
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
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
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1215
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   135
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   6885
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text positions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1905
      Left            =   90
      TabIndex        =   15
      Top             =   4950
      Width           =   5820
      Begin VB.HScrollBar HScroll5 
         Height          =   240
         Left            =   1620
         Max             =   10
         Min             =   -10
         TabIndex        =   24
         Top             =   1530
         Value           =   2
         Width           =   1500
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   240
         Left            =   1620
         Max             =   10
         Min             =   -10
         TabIndex        =   23
         Top             =   1170
         Value           =   2
         Width           =   1500
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   240
         LargeChange     =   10
         Left            =   1620
         Max             =   600
         Min             =   -300
         TabIndex        =   22
         Top             =   810
         Width           =   2220
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   240
         LargeChange     =   10
         Left            =   1620
         Max             =   600
         Min             =   -300
         TabIndex        =   20
         Top             =   450
         Width           =   2220
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   8
         Left            =   5310
         Picture         =   "CDC2.frx":0000
         Top             =   1170
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   7
         Left            =   4905
         Picture         =   "CDC2.frx":03A5
         Top             =   1170
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   6
         Left            =   4500
         Picture         =   "CDC2.frx":0745
         Top             =   1170
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   5
         Left            =   5310
         Picture         =   "CDC2.frx":0AE8
         Top             =   765
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   4
         Left            =   4905
         Picture         =   "CDC2.frx":0E86
         Top             =   765
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   3
         Left            =   4500
         Picture         =   "CDC2.frx":1231
         Top             =   765
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   2
         Left            =   5310
         Picture         =   "CDC2.frx":15D4
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   4905
         Picture         =   "CDC2.frx":1974
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   0
         Left            =   4500
         Picture         =   "CDC2.frx":1D15
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3195
         TabIndex        =   27
         Top             =   1530
         Width           =   420
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3195
         TabIndex        =   26
         Top             =   1170
         Width           =   420
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3870
         TabIndex        =   25
         Top             =   810
         Width           =   420
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3870
         TabIndex        =   21
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shadow Y"
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
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   1485
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shadow X"
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
         Left            =   135
         TabIndex        =   18
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y Position"
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
         Left            =   135
         TabIndex        =   17
         Top             =   765
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X Position"
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
         Left            =   135
         TabIndex        =   16
         Top             =   405
         Width           =   1410
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   495
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CDC2.frx":20B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CDC2.frx":2212
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CDC2.frx":236C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1635
      Left            =   90
      TabIndex        =   3
      Top             =   3285
      Width           =   5820
      Begin VB.CommandButton Command2 
         Caption         =   "Show fonts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   1215
         Width           =   1410
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Shadow"
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1170
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Text"
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   765
         Width           =   1005
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   810
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kBold"
               Object.ToolTipText     =   "Set text in bold"
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kItalic"
               Object.ToolTipText     =   "Set text in italic"
               ImageIndex      =   2
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "kUnderline"
               Object.ToolTipText     =   "Set text in underline"
               ImageIndex      =   3
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   10
         Left            =   3645
         Max             =   500
         Min             =   8
         TabIndex        =   5
         Top             =   360
         Value           =   8
         Width           =   1590
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shadowcolor"
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
         Left            =   3015
         TabIndex        =   12
         Top             =   1215
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Textcolor"
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
         Left            =   3015
         TabIndex        =   11
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
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
         Left            =   3015
         TabIndex        =   7
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5310
         TabIndex        =   6
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   135
      TabIndex        =   1
      Top             =   585
      Width           =   5820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done !"
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
      Left            =   4770
      TabIndex        =   0
      Top             =   7290
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   26
      Left            =   4275
      Picture         =   "CDC2.frx":24C6
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   25
      Left            =   3870
      Picture         =   "CDC2.frx":2867
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   24
      Left            =   3465
      Picture         =   "CDC2.frx":2C0D
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   23
      Left            =   4275
      Picture         =   "CDC2.frx":2FB0
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   22
      Left            =   3870
      Picture         =   "CDC2.frx":3350
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   21
      Left            =   3465
      Picture         =   "CDC2.frx":36F7
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   20
      Left            =   4275
      Picture         =   "CDC2.frx":3A98
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   19
      Left            =   3870
      Picture         =   "CDC2.frx":3E35
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   18
      Left            =   3465
      Picture         =   "CDC2.frx":41DA
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   17
      Left            =   5625
      Picture         =   "CDC2.frx":457F
      Top             =   810
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   16
      Left            =   5220
      Picture         =   "CDC2.frx":4924
      Top             =   810
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   15
      Left            =   4815
      Picture         =   "CDC2.frx":4CC4
      Top             =   765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   14
      Left            =   5625
      Picture         =   "CDC2.frx":5067
      Top             =   405
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   13
      Left            =   5220
      Picture         =   "CDC2.frx":5405
      Top             =   405
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   12
      Left            =   4815
      Picture         =   "CDC2.frx":57B0
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   11
      Left            =   5625
      Picture         =   "CDC2.frx":5B53
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   10
      Left            =   5220
      Picture         =   "CDC2.frx":5EF3
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   9
      Left            =   4815
      Picture         =   "CDC2.frx":6294
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   3105
      TabIndex        =   2
      Top             =   135
      Width           =   2625
   End
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuCol1 
         Caption         =   "Equalise textcolors"
      End
      Begin VB.Menu mnuCol2 
         Caption         =   "Equalise shadowcolors"
      End
      Begin VB.Menu mnuCol3 
         Caption         =   "Swith textcolors and shadowcolors"
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCol4 
         Caption         =   "Brighten textcolors"
      End
      Begin VB.Menu mnuCol5 
         Caption         =   "Darken textcolors"
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCol6 
         Caption         =   "Brighten shadowcolors"
      End
      Begin VB.Menu mnuCol7 
         Caption         =   "Darken shadowcolors"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCol8 
         Caption         =   "Brighten red text component"
      End
      Begin VB.Menu mnuCol9 
         Caption         =   "Brighten green text component"
      End
      Begin VB.Menu mnuCol10 
         Caption         =   "Brighten blue text component"
      End
      Begin VB.Menu mnuCol11 
         Caption         =   "Brighten red shadow component"
      End
      Begin VB.Menu mnuCol12 
         Caption         =   "Brighten green shadow component"
      End
      Begin VB.Menu mnuCol13 
         Caption         =   "Brighten blue shadow component"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCol14 
         Caption         =   "Darken red text component"
      End
      Begin VB.Menu mnuCol15 
         Caption         =   "Darken green text component"
      End
      Begin VB.Menu mnuCol16 
         Caption         =   "Darken blue text component"
      End
      Begin VB.Menu mnuCol17 
         Caption         =   "Darken red shadow component"
      End
      Begin VB.Menu mnuCol18 
         Caption         =   "Darken green shadow component"
      End
      Begin VB.Menu mnuCol19 
         Caption         =   "Darken blue shadow component"
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "Text positions"
      Begin VB.Menu mnuT0 
         Caption         =   "Equalise left aligment"
      End
      Begin VB.Menu mnuT1 
         Caption         =   "Equalise right alignment"
      End
      Begin VB.Menu mnuT2 
         Caption         =   "Center text positions"
      End
      Begin VB.Menu Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT3 
         Caption         =   "Make same distance 15 pixels"
      End
      Begin VB.Menu mnuT4 
         Caption         =   "Make same distance 20 pixels"
      End
      Begin VB.Menu mnuT5 
         Caption         =   "Make same distance 25 pixels"
      End
      Begin VB.Menu mnuT6 
         Caption         =   "Make same distance 30 pixels"
      End
      Begin VB.Menu Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT13 
         Caption         =   "Move all text 5 pixels right"
      End
      Begin VB.Menu mnuT14 
         Caption         =   "Move all text 10 pixels right"
      End
      Begin VB.Menu mnuT15 
         Caption         =   "Move all text 20 pixels right"
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT16 
         Caption         =   "Move all text 5 pixels left"
      End
      Begin VB.Menu mnuT17 
         Caption         =   "Move all text 10 pixels left"
      End
      Begin VB.Menu mnuT18 
         Caption         =   "Move all text 20 pixels left"
      End
      Begin VB.Menu bar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT7 
         Caption         =   "Move all  text 5 pixels down"
      End
      Begin VB.Menu mnuT8 
         Caption         =   "Move all text 10 pixels down"
      End
      Begin VB.Menu mnuT9 
         Caption         =   "Move all text 20 pixels down"
      End
      Begin VB.Menu Bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT10 
         Caption         =   "Move all text 5 pixels up"
      End
      Begin VB.Menu mnuT11 
         Caption         =   "Move all text 10 pixels up"
      End
      Begin VB.Menu mnuT12 
         Caption         =   "Move all text 20 pixels up"
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Fonts"
      Begin VB.Menu mnuF0 
         Caption         =   "Equal fonts"
      End
      Begin VB.Menu mnuF1 
         Caption         =   "Equal fontsizes"
      End
      Begin VB.Menu Bar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF2 
         Caption         =   "Increase fontsize by 2"
      End
      Begin VB.Menu mnuF3 
         Caption         =   "Decrease fontsize by 2"
      End
      Begin VB.Menu Bar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF4 
         Caption         =   "Make all text bold"
      End
      Begin VB.Menu mnuF5 
         Caption         =   "Make all text italic"
      End
      Begin VB.Menu mnuF6 
         Caption         =   "Make all text underline"
      End
      Begin VB.Menu Bar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF7 
         Caption         =   "Make all text not bold"
      End
      Begin VB.Menu mnuF8 
         Caption         =   "Make all text not italic"
      End
      Begin VB.Menu mnuF9 
         Caption         =   "Make all text not underline"
      End
   End
   Begin VB.Menu mnuCopyText 
      Caption         =   "Text"
      Begin VB.Menu mnuCopyF 
         Caption         =   "Copy to frontside"
      End
      Begin VB.Menu mnuCopyB 
         Caption         =   "Copy to backside"
      End
      Begin VB.Menu bar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClrTxt 
         Caption         =   "Clear all text"
      End
   End
End
Attribute VB_Name = "CDC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If PicIdx = 0 Then
    If Check1.Value = 0 Then
    CDC1.Lab0(q).Visible = False
    Else
    CDC1.Lab0(q).Visible = True
    End If
End If
If PicIdx = 1 Then
    If Check1.Value = 0 Then
    CDC1.Lab1(q).Visible = False
    Else
    CDC1.Lab1(q).Visible = True
    End If
End If
End Sub

Private Sub Check2_Click()
If PicIdx = 0 Then
    If Check2.Value = 0 Then
    CDC1.SLab0(q).Visible = False
    Else
    CDC1.SLab0(q).Visible = True
    End If
End If
If PicIdx = 1 Then
    If Check2.Value = 0 Then
    CDC1.SLab1(q).Visible = False
    Else
    CDC1.SLab1(q).Visible = True
    End If
End If
End Sub

Private Sub Combo1_Click()
If PicIdx = 0 Then CDC1.Lab0(q).Font = Combo1.Text
If PicIdx = 1 Then CDC1.Lab1(q).Font = Combo1.Text
If PicIdx = 0 Then CDC1.SLab0(q).Font = Combo1.Text
If PicIdx = 1 Then CDC1.SLab1(q).Font = Combo1.Text
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub Command1_Click()
CDC2.Hide
End Sub

Private Sub Command2_Click()
Showfont = 0
CDC10.Show 1
End Sub

Private Sub Form_Activate()
SetText
Text1.SetFocus
Toolbar2.Buttons(PicIdx + 1).Value = tbrPressed
Setmenus
End Sub

Private Sub Form_Load()
Me.Caption = CDCTitle
Me.Move 0, CDC1.Top, 6150, CDC1.Height - 10
Toolbar2.ImageList = CDC1.ImageList1
Toolbar2.Buttons(1).Image = 2
Toolbar2.Buttons(2).Image = 3
'T3D CDC2, Label1, 3, T3dRaiseInset
End Sub

Private Sub HScroll1_Change()
Label2 = Format(HScroll1.Value, "000")
If PicIdx = 0 Then
    CDC1.Lab0(q).FontSize = HScroll1.Value
    CDC1.SLab0(q).FontSize = HScroll1.Value
End If
If PicIdx = 1 Then
    CDC1.Lab1(q).FontSize = HScroll1.Value
    CDC1.SLab1(q).FontSize = HScroll1.Value
End If
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub HScroll2_Change()
Label9 = Format(HScroll2.Value, "000")
If PicIdx = 0 Then
    CDC1.Lab0(q).Left = HScroll2.Value
    CDC1.SLab0(q).Left = HScroll2.Value + ShX(0, q)
End If
If PicIdx = 1 Then
    CDC1.Lab1(q).Left = HScroll2.Value
    CDC1.SLab1(q).Left = HScroll2.Value + ShX(1, q)
End If
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub HScroll3_Change()
Label10 = Format(HScroll3.Value, "000")
If PicIdx = 0 Then
    CDC1.Lab0(q).Top = HScroll3.Value
    CDC1.SLab0(q).Top = HScroll3.Value + ShY(0, q)
End If
If PicIdx = 1 Then
    CDC1.Lab1(q).Top = HScroll3.Value
    CDC1.SLab1(q).Top = HScroll3.Value + ShY(1, q)
End If
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub HScroll4_Change()
Label11 = Format(HScroll4.Value, "00")
ShX(PicIdx, q) = HScroll4.Value
If PicIdx = 0 Then
CDC1.SLab0(q).Left = CDC1.Lab0(q).Left + ShX(0, q)
End If
If PicIdx = 1 Then
CDC1.SLab1(q).Left = CDC1.Lab1(q).Left + ShX(1, q)
End If
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub HScroll5_Change()
Label12 = Format(HScroll5.Value, "00")
ShY(PicIdx, q) = HScroll5.Value
If PicIdx = 0 Then
CDC1.SLab0(q).Top = CDC1.Lab0(q).Top + ShY(0, q)
End If
If PicIdx = 1 Then
CDC1.SLab1(q).Top = CDC1.Lab1(q).Top + ShY(1, q)
End If
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(Index) = Image1(Index + 18)
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image1(Index) = Image1(Index + 9)
Select Case Index
Case 0 'Top-Left
        HScroll2 = 0
        HScroll3 = 0
Case 1 'Top-Center
        HScroll3 = 0
        If PicIdx = 0 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width) / 2
        End If
        If PicIdx = 1 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width) / 2
        End If
Case 2 'Top-Right
        HScroll3 = 0
        If PicIdx = 0 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width
        End If
        If PicIdx = 1 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width
        End If
Case 3 'Center-Left
        HScroll2 = 0
        If PicIdx = 0 Then
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab0(q).Height) / 2
        End If
        If PicIdx = 1 Then
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab1(q).Height) / 2
        End If
Case 4 'Center-center
        If PicIdx = 0 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width) / 2
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab0(q).Height) / 2
        End If
        If PicIdx = 1 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width) / 2
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab1(q).Height) / 2
        End If
Case 5 'center-right
        If PicIdx = 0 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab0(q).Height) / 2
        End If
        If PicIdx = 1 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width
        HScroll3 = (CDC1.Pic1(PicIdx).Height - CDC1.SLab1(q).Height) / 2
       End If
Case 6 'bottom-left
        HScroll2 = 0
        If PicIdx = 0 Then
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab0(q).Height
        End If
        If PicIdx = 1 Then
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab1(q).Height
        End If
Case 7 'bottom-center
        If PicIdx = 0 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width) / 2
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab0(q).Height
        End If
        If PicIdx = 1 Then
        HScroll2 = (CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width) / 2
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab1(q).Height
        End If
Case 8 'bottom-right
        If PicIdx = 0 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab0(q).Width
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab0(q).Height
        End If
        If PicIdx = 1 Then
        HScroll2 = CDC1.Pic1(PicIdx).Width - CDC1.Lab1(q).Width
        HScroll3 = CDC1.Pic1(PicIdx).Height - CDC1.Lab1(q).Height
        End If
End Select
End Sub

Private Sub Label6_Click()
CD1.CancelError = False
CD1.Flags = 3
CD1.Color = Label6.BackColor
CD1.ShowColor
Label6.BackColor = CD1.Color
    If PicIdx = 0 Then CDC1.Lab0(q).ForeColor = CD1.Color
    If PicIdx = 1 Then CDC1.Lab1(q).ForeColor = CD1.Color
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub Label7_Click()
CD1.CancelError = False
CD1.Flags = 3
CD1.Color = Label7.BackColor
CD1.ShowColor
Label7.BackColor = CD1.Color
    If PicIdx = 0 Then CDC1.SLab0(q).ForeColor = CD1.Color
    If PicIdx = 1 Then CDC1.SLab1(q).ForeColor = CD1.Color
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub List1_Click()
q = List1.ListIndex
Select Case PicIdx
Case 0
DisplayTextSettings CDC1.Lab0(q), CDC1.SLab0(q)
Case 1
DisplayTextSettings CDC1.Lab1(q), CDC1.SLab1(q)
End Select
End Sub

Private Sub DisplayTextSettings(Ob1 As Object, Ob2 As Object)
On Error Resume Next
Text1 = List1.List(List1.ListIndex)
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Text1.SetFocus
Combo1.Text = Ob1.Font
HScroll1.Value = Ob1.FontSize
If Ob1.Visible = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If Ob2.Visible = True Then
Check2.Value = 1
Else
Check2.Value = 0
End If
Label6.BackColor = Ob1.ForeColor
Label7.BackColor = Ob2.ForeColor
Toolbar1.Buttons(1).Value = tbrUnpressed 'bold
Toolbar1.Buttons(2).Value = tbrUnpressed 'italic
Toolbar1.Buttons(3).Value = tbrUnpressed 'underline
If Ob1.FontBold = True Then Toolbar1.Buttons(1).Value = tbrPressed 'bold
If Ob1.FontItalic = True Then Toolbar1.Buttons(2).Value = tbrPressed 'bold
If Ob1.FontUnderline = True Then Toolbar1.Buttons(3).Value = tbrPressed 'bold
HScroll2.Value = Ob1.Left
HScroll3.Value = Ob1.Top
HScroll4.Value = ShX(PicIdx, q)
HScroll5.Value = ShY(PicIdx, q)
End Sub

Private Sub mnuClrTxt_Click()
Temp = MsgBox("This will clear all text" & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
With CDC1
For xx = 0 To 49
If PicIdx = 0 Then
    .Lab0(xx).Visible = False
    .Lab0(xx).Font = "Arial"
    .Lab0(xx).FontSize = 12
    .Lab0(xx).FontBold = False
    .Lab0(xx).FontItalic = False
    .Lab0(xx).FontUnderline = False
    .Lab0(xx) = "Front" & Format(xx, "00")
    .Lab0(xx).ForeColor = &HFFFF00
    .Lab0(xx).Left = 10
    .Lab0(xx).Top = 20 * xx
    ShX(0, xx) = 2
    ShY(0, xx) = 2
    .SLab0(xx).Visible = False
    .SLab0(xx).Font = "Arial"
    .SLab0(xx).FontSize = 12
    .SLab0(xx).FontBold = False
    .SLab0(xx).FontItalic = False
    .SLab0(xx).FontUnderline = False
    .SLab0(xx) = "Front" & Format(xx, "00")
    .SLab0(xx).ForeColor = 0
    .SLab0(xx).Left = .Lab0(xx).Left + ShX(0, xx)
    .SLab0(xx).Top = .Lab0(xx).Top + ShY(0, xx)
    Set .Lab0(xx).Container = .Pic1(0)
    Set .SLab0(xx).Container = .Pic1(0)
End If
If PicIdx = 1 Then
    .Lab1(xx).Visible = False
    .Lab1(xx).Font = "Arial"
    .Lab1(xx).FontSize = 12
    .Lab1(xx).FontBold = False
    .Lab1(xx).FontItalic = False
    .Lab1(xx).FontUnderline = False
    .Lab1(xx) = "Back" & Format(xx, "00")
    .Lab1(xx).ForeColor = &HFFFF00
    .Lab1(xx).Left = 10
    .Lab1(xx).Top = 20 * xx
    ShX(1, xx) = 2
    ShY(1, xx) = 2
    .SLab1(xx).Visible = False
    .SLab1(xx).Font = "Arial"
    .SLab1(xx).FontSize = 12
    .SLab1(xx).FontBold = False
    .SLab1(xx).FontItalic = False
    .SLab1(xx).FontUnderline = False
    .SLab1(xx) = "Back" & Format(xx, "00")
    .SLab1(xx).ForeColor = 0
    .SLab1(xx).Left = .Lab1(xx).Left + ShX(1, xx)
    .SLab1(xx).Top = .Lab1(xx).Top + ShY(1, xx)
    Set .Lab1(xx).Container = .Pic1(1)
    Set .SLab1(xx).Container = .Pic1(1)
End If
SetText
Next xx
End With
End Sub

Private Sub mnuCol1_Click() 'equalise visual textcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then CDC1.Lab0(xx).ForeColor = CDC1.Lab0(q).ForeColor
If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then CDC1.Lab1(xx).ForeColor = CDC1.Lab1(q).ForeColor
Next xx
End Sub

Private Sub mnuCol10_Click() 'brighten blue text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    B = B + 64
    If B > 255 Then B = 255
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If
If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    B = B + 64
    If B > 255 Then B = 255
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol11_Click() 'brighten red shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol12_Click() 'brighten green shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    G = G + 64
    If G > 255 Then G = 255
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    G = G + 64
    If G > 255 Then G = 255
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol13_Click() 'brighten blue shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    B = B + 64
    If B > 255 Then B = 255
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    B = B + 64
    If B > 255 Then B = 255
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol14_Click() 'darken red text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol15_Click() 'darken green text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    G = G - 64
    If G < 0 Then G = 0
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    G = G - 64
    If G < 0 Then G = 0
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol16_Click() 'darken blue text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    B = B - 64
    If B < 0 Then B = 0
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    B = B - 64
    If B < 0 Then B = 0
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol17_Click() 'darken red shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol18_Click() 'darken green shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    G = G - 64
    If G < 0 Then G = 0
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    G = G - 64
    If G < 0 Then G = 0
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol19_Click() 'darken blue shadow component
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    B = B - 64
    If B < 0 Then B = 0
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    B = B - 64
    If B < 0 Then B = 0
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol2_Click() 'equalise visible shadowcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then CDC1.SLab0(xx).ForeColor = CDC1.SLab0(q).ForeColor
If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then CDC1.SLab1(xx).ForeColor = CDC1.SLab1(q).ForeColor
Next xx
End Sub

Private Sub mnuCol3_Click() 'switch all colors
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    z = CDC1.Lab0(xx).ForeColor
    CDC1.Lab0(xx).ForeColor = CDC1.SLab0(xx).ForeColor
    CDC1.SLab0(xx).ForeColor = z
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    z = CDC1.Lab1(xx).ForeColor
    CDC1.Lab1(xx).ForeColor = CDC1.SLab1(xx).ForeColor
    CDC1.SLab1(xx).ForeColor = z
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol4_Click() 'brighten textcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    G = G + 64
    If G > 255 Then G = 255
    B = B + 64
    If B > 255 Then B = 255
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    G = G + 64
    If G > 255 Then G = 255
    B = B + 64
    If B > 255 Then B = 255
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol5_Click() 'darken textcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    G = G - 64
    If G < 0 Then G = 0
    B = B - 64
    If B < 0 Then B = 0
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    G = G - 64
    If G < 0 Then G = 0
    B = B - 64
    If B < 0 Then B = 0
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol6_Click() 'brighten shadowcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    G = G + 64
    If G > 255 Then G = 255
    B = B + 64
    If B > 255 Then B = 255
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    G = G + 64
    If G > 255 Then G = 255
    B = B + 64
    If B > 255 Then B = 255
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol7_Click() 'darken shadowcolors
For xx = 0 To 49
If PicIdx = 0 And CDC1.SLab0(xx).Visible = True Then
    R = CDC1.SLab0(xx).ForeColor Mod 256&
    G = ((CDC1.SLab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab0(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    G = G - 64
    If G < 0 Then G = 0
    B = B - 64
    If B < 0 Then B = 0
    CDC1.SLab0(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.SLab1(xx).Visible = True Then
    R = CDC1.SLab1(xx).ForeColor Mod 256&
    G = ((CDC1.SLab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.SLab1(xx).ForeColor And &HFF0000) / 65536
    R = R - 64
    If R < 0 Then R = 0
    G = G - 64
    If G < 0 Then G = 0
    B = B - 64
    If B < 0 Then B = 0
    CDC1.SLab1(xx).ForeColor = RGB(R, G, B)
    Label7.BackColor = CDC1.SLab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol8_Click() 'brighten red text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    R = R + 64
    If R > 255 Then R = 255
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCol9_Click() 'brighten green text component
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    R = CDC1.Lab0(xx).ForeColor Mod 256&
    G = ((CDC1.Lab0(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab0(xx).ForeColor And &HFF0000) / 65536
    G = G + 64
    If G > 255 Then G = 255
    CDC1.Lab0(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab0(xx).ForeColor
End If

If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    R = CDC1.Lab1(xx).ForeColor Mod 256&
    G = ((CDC1.Lab1(xx).ForeColor And &HFF00) / 256&) Mod 256&
    B = (CDC1.Lab1(xx).ForeColor And &HFF0000) / 65536
    G = G + 64
    If G > 255 Then G = 255
    CDC1.Lab1(xx).ForeColor = RGB(R, G, B)
    Label6.BackColor = CDC1.Lab1(xx).ForeColor
End If
Next xx
End Sub

Private Sub mnuCopyB_Click() 'copy to back
With CDC1
If PicIdx = 0 Then 'copy front to back
        CPText .Lab1, .Lab0, .SLab1, .SLab0, 1, 0
End If
End With
End Sub

Private Sub mnuCopyF_Click() 'copy to front
With CDC1
If PicIdx = 1 Then 'copy back to front
        CPText .Lab0, .Lab1, .SLab0, .SLab1, 0, 1
End If
End With
End Sub

Private Sub mnuCopyI_Click() 'copy to inlay
With CDC1
If PicIdx = 0 Then 'copy front to inlay
        CPText .Lab2, .Lab0, .SLab2, .SLab0, 2, 0
End If
If PicIdx = 1 Then 'copy back to inlay
        CPText .Lab2, .Lab1, .SLab2, .SLab1, 2, 1
End If
If PicIdx = 3 Then 'copy label to inlay
        CPText .Lab2, .Lab3, .SLab2, .SLab3, 2, 3
End If
End With
End Sub

Private Sub mnuCopyL_Click() 'copy to label
With CDC1
If PicIdx = 0 Then 'copy front to label
        CPText .Lab3, .Lab0, .SLab3, .SLab0, 3, 0
End If
If PicIdx = 1 Then 'copy back to label
        CPText .Lab3, .Lab1, .SLab3, .SLab1, 3, 1
End If
If PicIdx = 3 Then 'copy inlay to label
        CPText .Lab3, .Lab2, .SLab3, .SLab2, 3, 2
End If
End With
End Sub

Private Sub mnuF0_Click() 'equal fonts
Select Case PicIdx
Case 0
    For xx = 0 To 49
        If CDC1.Lab0(xx).Visible = True Then
        CDC1.Lab0(xx).Font = CDC1.Lab0(q).Font
        CDC1.SLab0(xx).Font = CDC1.Lab0(q).Font
        End If
    Next xx
Case 1
    For xx = 0 To 49
        If CDC1.Lab1(xx).Visible = True Then
        CDC1.Lab1(xx).Font = CDC1.Lab1(q).Font
        CDC1.SLab1(xx).Font = CDC1.Lab1(q).Font
        End If
    Next xx
End Select
End Sub

Private Sub mnuF1_Click() 'equal fontsizes
Select Case PicIdx
Case 0
    For xx = 0 To 49
        If CDC1.Lab0(xx).Visible = True Then
        CDC1.Lab0(xx).FontSize = CDC1.Lab0(q).FontSize
        CDC1.SLab0(xx).FontSize = CDC1.Lab0(q).FontSize
        End If
    Next xx
Case 1
    For xx = 0 To 49
        If CDC1.Lab1(xx).Visible = True Then
        CDC1.Lab1(xx).FontSize = CDC1.Lab1(q).FontSize
        CDC1.SLab1(xx).FontSize = CDC1.Lab1(q).FontSize
        End If
    Next xx
End Select
End Sub

Private Sub mnuF2_Click() 'increase fontsize  2
Select Case PicIdx
Case 0
    For xx = 0 To 49
        If CDC1.Lab0(xx).Visible = True Then
        CDC1.Lab0(xx).FontSize = CDC1.Lab0(xx).FontSize + 2
        CDC1.SLab0(xx).FontSize = CDC1.SLab0(xx).FontSize + 2
        End If
    Next xx
    HScroll1 = CDC1.Lab0(q).FontSize
Case 1
    For xx = 0 To 49
        If CDC1.Lab1(xx).Visible = True Then
        CDC1.Lab1(xx).FontSize = CDC1.Lab1(xx).FontSize + 2
        CDC1.SLab1(xx).FontSize = CDC1.SLab1(xx).FontSize + 2
        End If
    Next xx
    HScroll1 = CDC1.Lab1(q).FontSize
End Select
End Sub

Private Sub mnuF3_Click() 'decrease fontsize 2
Select Case PicIdx
Case 0
    For xx = 0 To 49
        If CDC1.Lab0(xx).Visible = True And CDC1.Lab0(xx).FontSize > 9 Then
        CDC1.Lab0(xx).FontSize = CDC1.Lab0(xx).FontSize - 2
        CDC1.SLab0(xx).FontSize = CDC1.SLab0(xx).FontSize - 2
        End If
    Next xx
    HScroll1 = CDC1.Lab0(q).FontSize
Case 1
    For xx = 0 To 49
        If CDC1.Lab1(xx).Visible = True And CDC1.Lab1(xx).FontSize > 9 Then
        CDC1.Lab1(xx).FontSize = CDC1.Lab1(xx).FontSize - 2
        CDC1.SLab1(xx).FontSize = CDC1.SLab1(xx).FontSize - 2
        End If
    Next xx
    HScroll1 = CDC1.Lab1(q).FontSize
End Select
End Sub

Private Sub mnuF4_Click() 'make bold
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontBold = True
    CDC1.SLab0(xx).FontBold = True
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(1).Value = tbrPressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontBold = True
    CDC1.SLab1(xx).FontBold = True
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(1).Value = tbrPressed
End Select
End Sub

Private Sub mnuF5_Click() 'make italic
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontItalic = True
    CDC1.SLab0(xx).FontItalic = True
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(2).Value = tbrPressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontItalic = True
    CDC1.SLab1(xx).FontItalic = True
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(2).Value = tbrPressed
End Select
End Sub

Private Sub mnuF6_Click() 'make underline
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontUnderline = True
    CDC1.SLab0(xx).FontUnderline = True
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(3).Value = tbrPressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontUnderline = True
    CDC1.SLab1(xx).FontUnderline = True
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(3).Value = tbrPressed
End Select
End Sub

Private Sub mnuF7_Click() 'make bold
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontBold = False
    CDC1.SLab0(xx).FontBold = False
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(1).Value = tbrUnpressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontBold = False
    CDC1.SLab1(xx).FontBold = False
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(1).Value = tbrUnpressed
End Select
End Sub

Private Sub mnuF8_Click() 'make italic
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontItalic = False
    CDC1.SLab0(xx).FontItalic = False
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(2).Value = tbrUnpressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontItalic = False
    CDC1.SLab1(xx).FontItalic = False
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(2).Value = tbrUnpressed
End Select
End Sub

Private Sub mnuF9_Click() 'make underline
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).FontUnderline = False
    CDC1.SLab0(xx).FontUnderline = False
    End If
Next xx
If CDC1.Lab0(q).Visible = True Then Toolbar1.Buttons(3).Value = tbrUnpressed
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).FontUnderline = False
    CDC1.SLab1(xx).FontUnderline = False
    End If
Next xx
If CDC1.Lab1(q).Visible = True Then Toolbar1.Buttons(3).Value = tbrUnpressed
End Select
End Sub

Private Sub mnuT0_Click() 'equalise left alignment
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).Left = CDC1.Lab0(q).Left
    CDC1.SLab0(xx).Left = CDC1.Lab0(q).Left + ShX(0, xx)
End If
If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).Left = CDC1.Lab1(q).Left
    CDC1.SLab1(xx).Left = CDC1.Lab1(q).Left + ShX(1, xx)
End If
Next xx
End Sub

Private Sub mnuT1_Click() 'equalise right alignment
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).Left = CDC1.Lab0(q).Left + CDC1.Lab0(q).Width - CDC1.Lab0(xx).Width
    CDC1.SLab0(xx).Left = CDC1.Lab0(q).Left + CDC1.Lab0(q).Width - CDC1.Lab0(xx).Width + ShX(0, xx)
End If
If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).Left = CDC1.Lab1(q).Left + CDC1.Lab1(q).Width - CDC1.Lab1(xx).Width
    CDC1.SLab1(xx).Left = CDC1.Lab1(q).Left + CDC1.Lab1(q).Width - CDC1.Lab1(xx).Width + ShX(1, xx)
End If
Next xx
End Sub

Private Sub mnuT10_Click()
MoveY -5
End Sub

Private Sub mnuT11_Click()
MoveY -10
End Sub

Private Sub mnuT12_Click()
MoveY -20
End Sub

Private Sub mnuT13_Click() 'move 5 right
MoveX 5
End Sub

Private Sub mnuT14_Click()
MoveX 10
End Sub

Private Sub mnuT15_Click()
MoveX 20
End Sub

Private Sub mnuT16_Click()
MoveX -5
End Sub

Private Sub mnuT17_Click()
MoveX -10
End Sub

Private Sub mnuT18_Click()
MoveX -20
End Sub

Private Sub mnuT2_Click() 'center text positions
For xx = 0 To 49
If PicIdx = 0 And CDC1.Lab0(xx).Visible = True Then
    CDC1.Lab0(xx).Left = (CDC1.Pic1(0).Width - CDC1.Lab0(xx).Width) / 2
    CDC1.SLab0(xx).Left = (CDC1.Pic1(0).Width - CDC1.Lab0(xx).Width) / 2 + ShX(0, xx)
End If
If PicIdx = 1 And CDC1.Lab1(xx).Visible = True Then
    CDC1.Lab1(xx).Left = (CDC1.Pic1(1).Width - CDC1.Lab1(xx).Width) / 2
    CDC1.SLab1(xx).Left = (CDC1.Pic1(1).Width - CDC1.Lab1(xx).Width) / 2 + ShX(1, xx)
End If
Next xx
End Sub

Private Sub mnuT3_Click() ' make same distance 15 pixels
MakeSameDistance 15
End Sub

Private Sub mnuT4_Click() 'make same distance 20 pixels
MakeSameDistance 20
End Sub

Private Sub mnuT5_Click() 'make same distance 25 pixels
MakeSameDistance 25
End Sub

Private Sub mnuT6_Click() 'make same distance 30 pixels
MakeSameDistance 30
End Sub

Private Sub mnuT7_Click() ' move 5 pixels down
MoveY 5
End Sub

Private Sub mnuT8_Click() 'move 10 pixels down
MoveY 10
End Sub

Private Sub mnuT9_Click()
MoveY 20
End Sub

Private Sub Text1_Change()
List1.List(List1.ListIndex) = Text1
If PicIdx = 0 Then 'front
CDC1.Lab0(q) = Text1
CDC1.SLab0(q) = Text1
End If
If PicIdx = 1 Then 'back
CDC1.Lab1(q) = Text1
CDC1.SLab1(q) = Text1
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "kBold"
If PicIdx = 0 Then
    If Toolbar1.Buttons(1).Value = tbrPressed Then
    CDC1.Lab0(q).FontBold = True
    Else
    CDC1.Lab0(q).FontBold = False
    End If
    CDC1.SLab0(q).FontBold = CDC1.Lab0(q).FontBold
End If
If PicIdx = 1 Then
    If Toolbar1.Buttons(1).Value = tbrPressed Then
    CDC1.Lab1(q).FontBold = True
    Else
    CDC1.Lab1(q).FontBold = False
    End If
    CDC1.SLab1(q).FontBold = CDC1.Lab1(q).FontBold
End If

Case "kItalic"
If PicIdx = 0 Then
    If Toolbar1.Buttons(2).Value = tbrPressed Then
    CDC1.Lab0(q).FontItalic = True
    Else
    CDC1.Lab0(q).FontItalic = False
    End If
    CDC1.SLab0(q).FontItalic = CDC1.Lab0(q).FontItalic
End If
If PicIdx = 1 Then
    If Toolbar1.Buttons(2).Value = tbrPressed Then
    CDC1.Lab1(q).FontItalic = True
    Else
    CDC1.Lab1(q).FontItalic = False
    End If
    CDC1.SLab1(q).FontItalic = CDC1.Lab1(q).FontItalic
End If
Case "kUnderline"
If PicIdx = 0 Then
    If Toolbar1.Buttons(3).Value = tbrPressed Then
    CDC1.Lab0(q).FontUnderline = True
    Else
    CDC1.Lab0(q).FontUnderline = False
    End If
    CDC1.SLab0(q).FontUnderline = CDC1.Lab0(q).FontUnderline
End If
If PicIdx = 1 Then
    If Toolbar1.Buttons(3).Value = tbrPressed Then
    CDC1.Lab1(q).FontUnderline = True
    Else
    CDC1.Lab1(q).FontUnderline = False
    End If
    CDC1.SLab1(q).FontUnderline = CDC1.Lab1(q).FontUnderline
End If
End Select
CDC1.Pic1(PicIdx).Refresh
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "kFront"
CDC1.Pic1(0).Visible = True
CDC1.Pic1(1).Visible = False
PicIdx = 0
If Re(PicIdx) = True Then
CDC1.Toolbar1.Buttons(1).Enabled = True
Else
CDC1.Toolbar1.Buttons(1).Enabled = False
End If
CDC1.Label1 = "FRONTSIDE"
CDC1.Label3 = PicW & " X " & PicH
Case "kBack"
CDC1.Pic1(0).Visible = False
CDC1.Pic1(1).Visible = True
PicIdx = 1
If Re(PicIdx) = True Then
CDC1.Toolbar1.Buttons(1).Enabled = True
Else
CDC1.Toolbar1.Buttons(1).Enabled = False
End If
CDC1.Label1 = "BACKSIDE"
CDC1.Label3 = PicW & " X " & PicH
End Select
SetText
Setmenus
End Sub

Private Sub MakeSameDistance(Ds%)
Select Case PicIdx
Case 0
For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
    p = xx
    Exit For
    End If
Next xx
pY = 0
        For xx = p To 49
                If CDC1.Lab0(xx).Visible = True Then
                CDC1.Lab0(xx).Top = CDC1.Lab0(p).Top + pY * Ds
                CDC1.SLab0(xx).Top = CDC1.Lab0(p).Top + pY * Ds + ShY(0, xx)
               pY = pY + 1
                End If
        Next xx
    HScroll3 = CDC1.Lab0(q).Top
Case 1
For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
    p = xx
    Exit For
    End If
Next xx
pY = 0
        For xx = p To 49
                If CDC1.Lab1(xx).Visible = True Then
                CDC1.Lab1(xx).Top = CDC1.Lab1(p).Top + pY * Ds
                CDC1.SLab1(xx).Top = CDC1.Lab1(p).Top + pY * Ds + ShY(0, xx)
               pY = pY + 1
                End If
        Next xx
    HScroll3 = CDC1.Lab1(q).Top
End Select
End Sub

Private Sub MoveY(Ds%)
Select Case PicIdx
Case 0
    For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
        CDC1.Lab0(xx).Top = CDC1.Lab0(xx).Top + Ds
        CDC1.SLab0(xx).Top = CDC1.Lab0(xx).Top + ShY(0, xx)
    End If
    Next xx
    HScroll3 = CDC1.Lab0(q).Top
Case 1
    For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
        CDC1.Lab1(xx).Top = CDC1.Lab1(xx).Top + Ds
        CDC1.SLab1(xx).Top = CDC1.Lab1(xx).Top + ShY(1, xx)
    End If
    Next xx
    HScroll3 = CDC1.Lab1(q).Top
End Select
End Sub

Private Sub MoveX(Ds%)
Select Case PicIdx
Case 0
    For xx = 0 To 49
    If CDC1.Lab0(xx).Visible = True Then
        CDC1.Lab0(xx).Left = CDC1.Lab0(xx).Left + Ds
        CDC1.SLab0(xx).Left = CDC1.Lab0(xx).Left + ShX(0, xx)
    End If
    Next xx
    HScroll2 = CDC1.Lab0(q).Left
Case 1
    For xx = 0 To 49
    If CDC1.Lab1(xx).Visible = True Then
        CDC1.Lab1(xx).Left = CDC1.Lab1(xx).Left + Ds
        CDC1.SLab1(xx).Left = CDC1.Lab1(xx).Left + ShX(1, xx)
    End If
    Next xx
    HScroll2 = CDC1.Lab1(q).Left
End Select
End Sub

Private Sub Setmenus()
mnuCopyF.Enabled = True
mnuCopyB.Enabled = True
'mnuCopyI.Enabled = True
'mnuCopyL.Enabled = True
If PicIdx = 0 Then mnuCopyF.Enabled = False
If PicIdx = 1 Then mnuCopyB.Enabled = False
End Sub

Private Sub CPText(Ob1 As Object, Ob2 As Object, Ob3 As Object, Ob4 As Object, ArIdx1, ArIdx2)
For xx = 0 To 49
    If Ob2(xx).Visible = True Then
        Ob1(xx) = Ob2(xx)
        Ob1(xx).Visible = True
        Ob1(xx).Font = Ob2(xx).Font
        Ob1(xx).FontSize = Ob2(xx).FontSize
        Ob1(xx).ForeColor = Ob2(xx).ForeColor
        Ob1(xx).FontBold = Ob2(xx).FontBold
        Ob1(xx).FontItalic = Ob2(xx).FontItalic
        Ob1(xx).FontUnderline = Ob2(xx).FontUnderline
        Ob1(xx).Left = Ob2(xx).Left
        Ob1(xx).Top = Ob2(xx).Top
                    Ob3(xx) = Ob4(xx)
                    Ob3(xx).Visible = Ob4(xx).Visible
                    Ob3(xx).Font = Ob4(xx).Font
                    Ob3(xx).FontSize = Ob4(xx).FontSize
                    Ob3(xx).ForeColor = Ob4(xx).ForeColor
                    Ob3(xx).FontBold = Ob4(xx).FontBold
                    Ob3(xx).FontItalic = Ob4(xx).FontItalic
                    Ob3(xx).FontUnderline = Ob4(xx).FontUnderline
                    Ob3(xx).Left = Ob4(xx).Left
                    Ob3(xx).Top = Ob4(xx).Top
            ShX(ArIdx1, xx) = ShX(ArIdx2, xx)
            ShY(0, xx) = ShY(1, xx)
End If
Next xx
End Sub

