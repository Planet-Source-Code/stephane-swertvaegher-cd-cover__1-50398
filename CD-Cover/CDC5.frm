VERSION 5.00
Begin VB.Form CDC5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   135
      Width           =   5550
      Begin VB.OptionButton Option1 
         Caption         =   "Linear background #3"
         Height          =   240
         Index           =   3
         Left            =   1665
         TabIndex        =   18
         Top             =   1260
         Width           =   1905
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         Height          =   285
         Left            =   945
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   246
         TabIndex        =   15
         Top             =   2070
         Width           =   3750
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Vertical gradient"
         Height          =   240
         Left            =   1665
         TabIndex        =   10
         Top             =   1575
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linear background #2"
         Height          =   240
         Index           =   2
         Left            =   1665
         TabIndex        =   9
         Top             =   945
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linear background  #1"
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   8
         Top             =   630
         Width           =   1950
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Equal background"
         Height          =   240
         Index           =   0
         Left            =   1665
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   3480
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   13
         Top             =   2520
         Value           =   1
         Width           =   2805
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color 3"
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
         TabIndex        =   17
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   16
         Top             =   1170
         Width           =   375
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
         Left            =   4185
         TabIndex        =   14
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alphablend"
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
         Left            =   135
         TabIndex        =   12
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   6
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   5
         Top             =   450
         Width           =   375
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
         TabIndex        =   4
         Top             =   810
         Width           =   735
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
         TabIndex        =   3
         Top             =   450
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mix inlay sides with color"
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
      Height          =   2940
      Index           =   12
      Left            =   270
      TabIndex        =   182
      Top             =   90
      Width           =   5550
      Begin VB.OptionButton Option8 
         Caption         =   "Gradient color #3"
         Height          =   240
         Index           =   3
         Left            =   2700
         TabIndex        =   196
         Top             =   1395
         Width           =   1995
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Gradient color #2"
         Height          =   240
         Index           =   2
         Left            =   2700
         TabIndex        =   193
         Top             =   1080
         Width           =   1995
      End
      Begin VB.HScrollBar HScroll40 
         Height          =   240
         Left            =   1575
         Max             =   10
         Min             =   1
         TabIndex        =   191
         Top             =   1890
         Value           =   1
         Width           =   2130
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Gradient color #1"
         Height          =   240
         Index           =   1
         Left            =   2700
         TabIndex        =   189
         Top             =   810
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Fixed color"
         Height          =   240
         Index           =   0
         Left            =   2700
         TabIndex        =   188
         Top             =   540
         Width           =   1995
      End
      Begin VB.PictureBox Pic3 
         AutoRedraw      =   -1  'True
         Height          =   285
         Left            =   1800
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   126
         TabIndex        =   187
         Top             =   2430
         Width           =   1950
      End
      Begin VB.Label Label74 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1530
         TabIndex        =   195
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label72 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inlaycolor3:"
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
         Left            =   405
         TabIndex        =   194
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         TabIndex        =   192
         Top             =   1890
         Width           =   600
      End
      Begin VB.Label Label70 
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
         Left            =   405
         TabIndex        =   190
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label69 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1530
         TabIndex        =   186
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label72 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inlaycolor2:"
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
         Left            =   405
         TabIndex        =   185
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label71 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1530
         TabIndex        =   184
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label72 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inlaycolor1:"
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
         Left            =   405
         TabIndex        =   183
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Index           =   1
      Left            =   270
      TabIndex        =   19
      Top             =   135
      Width           =   5550
      Begin VB.HScrollBar HScroll6 
         Height          =   240
         Left            =   1305
         Max             =   50
         Min             =   1
         TabIndex        =   40
         Top             =   2520
         Value           =   1
         Width           =   2805
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   240
         LargeChange     =   10
         Left            =   1305
         Max             =   500
         Min             =   10
         TabIndex        =   38
         Top             =   2205
         Value           =   10
         Width           =   2805
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   240
         LargeChange     =   10
         Left            =   1305
         Max             =   500
         TabIndex        =   36
         Top             =   1890
         Value           =   10
         Width           =   2805
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   240
         LargeChange     =   10
         Left            =   1305
         Max             =   500
         TabIndex        =   34
         Top             =   1575
         Value           =   10
         Width           =   2805
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Circular gradient #1"
         Height          =   240
         Index           =   0
         Left            =   1665
         TabIndex        =   23
         Top             =   495
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Circular gradient #2"
         Height          =   240
         Index           =   1
         Left            =   1665
         TabIndex        =   22
         Top             =   810
         Width           =   1725
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   240
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   21
         Top             =   1260
         Value           =   1
         Width           =   2805
      End
      Begin VB.PictureBox Pic2 
         AutoRedraw      =   -1  'True
         Height          =   285
         Left            =   3465
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   126
         TabIndex        =   20
         Top             =   495
         Width           =   1950
      End
      Begin VB.Label Label22 
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
         Left            =   4185
         TabIndex        =   41
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label21 
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
         Left            =   4185
         TabIndex        =   39
         Top             =   2205
         Width           =   600
      End
      Begin VB.Label Label20 
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
         Left            =   4185
         TabIndex        =   37
         Top             =   1890
         Width           =   600
      End
      Begin VB.Label Label19 
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
         Left            =   4185
         TabIndex        =   35
         Top             =   1575
         Width           =   600
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aspect"
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
         Left            =   135
         TabIndex        =   33
         Top             =   2475
         Width           =   1140
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Radius "
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
         Left            =   135
         TabIndex        =   32
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Offset Y"
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
         Left            =   135
         TabIndex        =   31
         Top             =   1845
         Width           =   1140
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Offset X"
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
         Left            =   135
         TabIndex        =   30
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         TabIndex        =   29
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label15 
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
         TabIndex        =   28
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   27
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   26
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alphablend"
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
         Left            =   135
         TabIndex        =   25
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Left            =   4185
         TabIndex        =   24
         Top             =   1260
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add sinus lines"
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
      Height          =   2940
      Index           =   11
      Left            =   270
      TabIndex        =   160
      Top             =   90
      Width           =   5550
      Begin VB.HScrollBar HScroll39 
         Height          =   285
         Left            =   2745
         Max             =   25
         Min             =   1
         TabIndex        =   180
         Top             =   1710
         Value           =   1
         Width           =   1185
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Absolute reversed"
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
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   178
         Top             =   2565
         Width           =   2400
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Vertical"
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
         Height          =   195
         Left            =   3645
         TabIndex        =   177
         Top             =   2295
         Width           =   1590
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Absolute"
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
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   176
         Top             =   2295
         Width           =   1635
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Normal"
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
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   175
         Top             =   2070
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll38 
         Height          =   285
         Left            =   1440
         Max             =   10
         Min             =   1
         TabIndex        =   172
         Top             =   1305
         Value           =   1
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll37 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   99
         Min             =   1
         TabIndex        =   167
         Top             =   990
         Value           =   2
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll36 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   250
         Min             =   1
         TabIndex        =   166
         Top             =   675
         Value           =   2
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll35 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   250
         Min             =   1
         TabIndex        =   161
         Top             =   360
         Value           =   2
         Width           =   2445
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   181
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linewidth:"
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
         Index           =   12
         Left            =   1755
         TabIndex        =   179
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label Label53 
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
         Index           =   11
         Left            =   180
         TabIndex        =   174
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   173
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label66 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1125
         TabIndex        =   171
         Top             =   1710
         Width           =   375
      End
      Begin VB.Label Label65 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linecolor:"
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
         TabIndex        =   170
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   169
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   168
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wavelength:"
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
         Index           =   10
         Left            =   180
         TabIndex        =   165
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ampltude:"
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
         Index           =   9
         Left            =   180
         TabIndex        =   164
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   163
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance:"
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
         Index           =   8
         Left            =   180
         TabIndex        =   162
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add lines"
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
      Height          =   2940
      Index           =   2
      Left            =   270
      TabIndex        =   42
      Top             =   90
      Width           =   5550
      Begin VB.OptionButton Option3 
         Caption         =   "Crossed line 45°"
         Height          =   240
         Index           =   5
         Left            =   3825
         TabIndex        =   74
         Top             =   2205
         Width           =   1635
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Single line 225°"
         Height          =   240
         Index           =   4
         Left            =   3825
         TabIndex        =   73
         Top             =   1935
         Width           =   1635
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Single line 45°"
         Height          =   240
         Index           =   3
         Left            =   3825
         TabIndex        =   72
         Top             =   1665
         Width           =   1365
      End
      Begin VB.HScrollBar HScroll14 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   250
         Min             =   1
         TabIndex        =   70
         Top             =   2610
         Value           =   5
         Width           =   1950
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Crossed line"
         Height          =   240
         Index           =   2
         Left            =   3825
         TabIndex        =   68
         Top             =   1395
         Width           =   1365
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Single line Y"
         Height          =   240
         Index           =   1
         Left            =   3825
         TabIndex        =   67
         Top             =   1125
         Width           =   1365
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Single line X"
         Height          =   240
         Index           =   0
         Left            =   3825
         TabIndex        =   66
         Top             =   855
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.HScrollBar HScroll13 
         Height          =   240
         Left            =   1080
         Max             =   25
         Min             =   1
         TabIndex        =   64
         Top             =   2295
         Value           =   5
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll12 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   100
         Min             =   1
         TabIndex        =   57
         Top             =   1980
         Value           =   1
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll11 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   800
         Min             =   10
         TabIndex        =   56
         Top             =   1665
         Value           =   10
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll10 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   800
         Min             =   10
         TabIndex        =   55
         Top             =   1350
         Value           =   10
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   500
         TabIndex        =   54
         Top             =   1035
         Value           =   250
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   240
         LargeChange     =   10
         Left            =   1080
         Max             =   500
         TabIndex        =   53
         Top             =   720
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   240
         Left            =   3195
         Max             =   10
         Min             =   1
         TabIndex        =   46
         Top             =   405
         Value           =   1
         Width           =   1500
      End
      Begin VB.Label Label34 
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
         Left            =   3060
         TabIndex        =   71
         Top             =   2610
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance:"
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
         Index           =   6
         Left            =   135
         TabIndex        =   69
         Top             =   2610
         Width           =   915
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   3060
         TabIndex        =   65
         Top             =   2295
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linewidth:"
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
         Index           =   5
         Left            =   135
         TabIndex        =   63
         Top             =   2295
         Width           =   915
      End
      Begin VB.Label Label32 
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
         Left            =   3060
         TabIndex        =   62
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label31 
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
         Left            =   3060
         TabIndex        =   61
         Top             =   1665
         Width           =   645
      End
      Begin VB.Label Label30 
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
         Left            =   3060
         TabIndex        =   60
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label29 
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
         Left            =   3060
         TabIndex        =   59
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label28 
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
         Left            =   3060
         TabIndex        =   58
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number:"
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
         Index           =   4
         Left            =   135
         TabIndex        =   52
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Length Y:"
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
         TabIndex        =   51
         Top             =   1665
         Width           =   915
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Length X:"
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
         TabIndex        =   50
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label Label27 
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
         Index           =   1
         Left            =   135
         TabIndex        =   49
         Top             =   1035
         Width           =   915
      End
      Begin VB.Label Label27 
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
         Index           =   0
         Left            =   135
         TabIndex        =   48
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label26 
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
         Left            =   4725
         TabIndex        =   47
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label25 
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
         Left            =   1980
         TabIndex        =   45
         Top             =   405
         Width           =   1140
      End
      Begin VB.Label Label24 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1215
         TabIndex        =   44
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linecolor:"
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
         Height          =   330
         Left            =   180
         TabIndex        =   43
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mosaïc"
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
      Height          =   2940
      Index           =   10
      Left            =   270
      TabIndex        =   155
      Top             =   90
      Width           =   5550
      Begin VB.CheckBox Check6 
         Caption         =   "Blurred mosaïc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1845
         TabIndex        =   159
         Top             =   1170
         Width           =   1770
      End
      Begin VB.HScrollBar HScroll34 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   50
         Min             =   1
         TabIndex        =   156
         Top             =   675
         Value           =   2
         Width           =   2445
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance:"
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
         Index           =   5
         Left            =   180
         TabIndex        =   158
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   157
         Top             =   675
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add blinds"
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
      Height          =   2940
      Index           =   9
      Left            =   270
      TabIndex        =   147
      Top             =   90
      Width           =   5550
      Begin VB.CheckBox Check5 
         Caption         =   "Bumped blinds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2925
         TabIndex        =   154
         Top             =   2475
         Width           =   1635
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Vertical blinds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   153
         Top             =   2385
         Width           =   1860
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Horizontal blinds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   152
         Top             =   2070
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Reverse blinds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   151
         Top             =   2160
         Width           =   1725
      End
      Begin VB.HScrollBar HScroll33 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   250
         Min             =   2
         TabIndex        =   148
         Top             =   675
         Value           =   2
         Width           =   2445
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   150
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance:"
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
         Index           =   4
         Left            =   180
         TabIndex        =   149
         Top             =   675
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tyle"
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
      Height          =   2940
      Index           =   8
      Left            =   270
      TabIndex        =   140
      Top             =   90
      Width           =   5550
      Begin VB.HScrollBar HScroll32 
         Height          =   285
         LargeChange     =   10
         Left            =   1215
         Max             =   10
         Min             =   1
         TabIndex        =   146
         Top             =   1305
         Value           =   2
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll31 
         Height          =   285
         LargeChange     =   10
         Left            =   1215
         Max             =   10
         Min             =   1
         TabIndex        =   145
         Top             =   945
         Value           =   2
         Width           =   2445
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tile X:"
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
         Index           =   7
         Left            =   180
         TabIndex        =   144
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tile Y:"
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
         Index           =   6
         Left            =   180
         TabIndex        =   143
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   142
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   141
         Top             =   945
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Echo"
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
      Height          =   2940
      Index           =   7
      Left            =   270
      TabIndex        =   125
      Top             =   90
      Width           =   5550
      Begin VB.HScrollBar HScroll30 
         Height          =   240
         LargeChange     =   10
         Left            =   1215
         Max             =   100
         Min             =   -100
         TabIndex        =   135
         Top             =   1935
         Value           =   100
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll29 
         Height          =   240
         LargeChange     =   10
         Left            =   1215
         Max             =   100
         Min             =   -100
         TabIndex        =   134
         Top             =   1620
         Value           =   -100
         Width           =   2445
      End
      Begin VB.HScrollBar HScroll28 
         Height          =   240
         LargeChange     =   10
         Left            =   1215
         Max             =   49
         Min             =   1
         TabIndex        =   133
         Top             =   1305
         Value           =   49
         Width           =   2445
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Default"
         Height          =   285
         Left            =   270
         TabIndex        =   132
         Top             =   2430
         Width           =   1545
      End
      Begin VB.HScrollBar HScroll27 
         Height          =   240
         LargeChange     =   10
         Left            =   1215
         Max             =   75
         Min             =   1
         TabIndex        =   131
         Top             =   990
         Value           =   25
         Width           =   2445
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Phasing"
         Height          =   285
         Left            =   180
         TabIndex        =   126
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   139
         Top             =   1935
         Width           =   780
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   138
         Top             =   1620
         Width           =   780
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   137
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Height          =   285
         Left            =   3735
         TabIndex        =   136
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label53 
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
         Index           =   3
         Left            =   180
         TabIndex        =   130
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label53 
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
         Index           =   2
         Left            =   180
         TabIndex        =   129
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reduced:"
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
         TabIndex        =   128
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number:"
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
         TabIndex        =   127
         Top             =   945
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave picture X"
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
      Height          =   2940
      Index           =   5
      Left            =   270
      TabIndex        =   107
      Top             =   90
      Width           =   5550
      Begin VB.OptionButton Option4 
         Caption         =   "Absolute Sinus"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   115
         Top             =   1890
         Width           =   1995
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Sinus"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   114
         Top             =   1620
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.HScrollBar HScroll24 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   99
         Min             =   1
         TabIndex        =   112
         Top             =   900
         Value           =   1
         Width           =   2085
      End
      Begin VB.HScrollBar HScroll23 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   250
         Min             =   1
         TabIndex        =   110
         Top             =   540
         Value           =   1
         Width           =   2085
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   3555
         TabIndex        =   113
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label49 
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
         Left            =   3555
         TabIndex        =   111
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wavelength:"
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
         TabIndex        =   109
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude:"
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
         TabIndex        =   108
         Top             =   540
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contour"
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
      Height          =   1140
      Index           =   4
      Left            =   270
      TabIndex        =   104
      Top             =   1080
      Width           =   5550
      Begin VB.Label Label47 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2250
         TabIndex        =   106
         Top             =   540
         Width           =   465
      End
      Begin VB.Label Label46 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contourcolor:"
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
         Height          =   330
         Left            =   900
         TabIndex        =   105
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add circles"
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
      Height          =   2940
      Index           =   3
      Left            =   270
      TabIndex        =   75
      Top             =   90
      Width           =   5550
      Begin VB.CheckBox Check2 
         Caption         =   "Double aspect"
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
         Height          =   510
         Left            =   4365
         TabIndex        =   103
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Center"
         Height          =   285
         Left            =   4365
         TabIndex        =   102
         Top             =   900
         Width           =   1005
      End
      Begin VB.HScrollBar HScroll18 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   250
         Min             =   1
         TabIndex        =   100
         Top             =   1350
         Value           =   5
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll22 
         Height          =   240
         Left            =   3150
         Max             =   10
         Min             =   1
         TabIndex        =   82
         Top             =   360
         Value           =   1
         Width           =   1500
      End
      Begin VB.HScrollBar HScroll21 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   800
         TabIndex        =   81
         Top             =   1035
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll20 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   800
         TabIndex        =   80
         Top             =   720
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll19 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   50
         Min             =   1
         TabIndex        =   79
         Top             =   1665
         Value           =   1
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll17 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   100
         Min             =   1
         TabIndex        =   78
         Top             =   1980
         Value           =   1
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll16 
         Height          =   240
         Left            =   1530
         Max             =   25
         Min             =   1
         TabIndex        =   77
         Top             =   2295
         Value           =   5
         Width           =   1950
      End
      Begin VB.HScrollBar HScroll15 
         Height          =   240
         LargeChange     =   10
         Left            =   1530
         Max             =   250
         Min             =   1
         TabIndex        =   76
         Top             =   2610
         Value           =   5
         Width           =   1950
      End
      Begin VB.Label Label38 
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
         Left            =   3555
         TabIndex        =   101
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Base radius:"
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
         Index           =   10
         Left            =   360
         TabIndex        =   99
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linecolor:"
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
         Height          =   330
         Left            =   360
         TabIndex        =   98
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label44 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1395
         TabIndex        =   97
         Top             =   315
         Width           =   465
      End
      Begin VB.Label Label43 
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
         Left            =   1935
         TabIndex        =   96
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label42 
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
         Left            =   4680
         TabIndex        =   95
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label27 
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
         Index           =   13
         Left            =   360
         TabIndex        =   94
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label27 
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
         Index           =   12
         Left            =   360
         TabIndex        =   93
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aspect:"
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
         Index           =   11
         Left            =   360
         TabIndex        =   92
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number:"
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
         Index           =   9
         Left            =   360
         TabIndex        =   91
         Top             =   1980
         Width           =   1140
      End
      Begin VB.Label Label41 
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
         Left            =   3555
         TabIndex        =   90
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label40 
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
         Left            =   3555
         TabIndex        =   89
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label39 
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
         Left            =   3555
         TabIndex        =   88
         Top             =   1665
         Width           =   645
      End
      Begin VB.Label Label37 
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
         Left            =   3555
         TabIndex        =   87
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linewidth:"
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
         Index           =   8
         Left            =   360
         TabIndex        =   86
         Top             =   2295
         Width           =   1140
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   3555
         TabIndex        =   85
         Top             =   2295
         Width           =   645
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance:"
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
         Index           =   7
         Left            =   360
         TabIndex        =   84
         Top             =   2610
         Width           =   1140
      End
      Begin VB.Label Label35 
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
         Left            =   3555
         TabIndex        =   83
         Top             =   2610
         Width           =   645
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show me"
      Height          =   375
      Left            =   135
      TabIndex        =   11
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4635
      TabIndex        =   1
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3195
      TabIndex        =   0
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave picture Y"
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
      Height          =   2940
      Index           =   6
      Left            =   270
      TabIndex        =   116
      Top             =   90
      Width           =   5550
      Begin VB.HScrollBar HScroll26 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   250
         Min             =   1
         TabIndex        =   120
         Top             =   540
         Value           =   1
         Width           =   2085
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Sinus"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   118
         Top             =   1620
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Absolute Sinus"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   117
         Top             =   1890
         Width           =   1995
      End
      Begin VB.HScrollBar HScroll25 
         Height          =   285
         LargeChange     =   10
         Left            =   1440
         Max             =   99
         Min             =   1
         TabIndex        =   119
         Top             =   900
         Value           =   1
         Width           =   2085
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude:"
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
         Left            =   180
         TabIndex        =   124
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wavelength:"
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
         TabIndex        =   123
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label52 
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
         Left            =   3555
         TabIndex        =   122
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   3555
         TabIndex        =   121
         Top             =   900
         Width           =   780
      End
   End
End
Attribute VB_Name = "CDC5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'ok
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
If Cidx = 0 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    ChangeColor
End If
If Cidx = 1 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    Changecolor2
End If
If Cidx = 2 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    SetLines
End If
If Cidx = 3 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    SetCircles
End If
If Cidx = 4 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
GetPicData .Pic1(PicIdx)
Contour Label47.BackColor
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 5 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    GetPicData .Pic1(PicIdx)
        If Option4(0).Value = True Then
        EffectX .Pic1(PicIdx), HScroll23, HScroll24, 0
        Else
        EffectX .Pic1(PicIdx), HScroll23, HScroll24, 1
        End If
    SetPicData .Pic1(PicIdx)
    .Pic1(PicIdx).Refresh
End If
If Cidx = 6 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SaveRedo
    GetPicData .Pic1(PicIdx)
        If Option5(0).Value = True Then
        EffectY .Pic1(PicIdx), HScroll26, HScroll25, 0
        Else
        EffectY .Pic1(PicIdx), HScroll26, HScroll25, 1
        End If
    SetPicData .Pic1(PicIdx)
    .Pic1(PicIdx).Refresh
End If
If Cidx = 7 Then 'echo
.Pic3.Picture = Nothing
.Pic1(PicIdx) = Im
SaveRedo
.Pic3.Picture = .Pic1(PicIdx).Image
Echo .Pic1(PicIdx), HScroll27, HScroll28, HScroll29, HScroll30
End If
If Cidx = 8 Then 'tile
.Pic3.Picture = Nothing
.Pic1(PicIdx) = Im
SaveRedo
.Pic3.Picture = .Pic1(PicIdx).Image
.Pic3.Refresh
Tile .Pic1(PicIdx), HScroll31, HScroll32
End If
If Cidx = 9 Then 'blinds
    .Pic1(PicIdx) = Im
    SaveRedo
    GetPicData .Pic1(PicIdx)
    If Check4 = 0 And Option6(0) = True And Check5 = 0 Then 'hor. blinds
    Blinds HScroll33, False, False
    End If
    If Check4 = 1 And Option6(0) = True And Check5 = 0 Then 'hor. blinds reversed
    Blinds HScroll33, True, False
    End If
    If Check4 = 0 And Option6(1) = True And Check5 = 0 Then 'vert. blinds
    Blinds HScroll33, False, True
    End If
    If Check4 = 1 And Option6(1) = True And Check5 = 0 Then 'vert. blinds reversed
    Blinds HScroll33, True, True
    End If
    If Option6(0) = True And Check5 = 1 Then 'hor. bumped blinds
    Blinds3 HScroll33, False
    End If
    If Option6(1) = True And Check5 = 1 Then 'vert. bumped blinds
    Blinds3 HScroll33, True
    End If
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 10 Then 'mosaic
    .Pic1(PicIdx) = Im
    SaveRedo
    GetPicData .Pic1(PicIdx)
    If Check6 = 0 Then Mozaic HScroll34
    If Check6 = 1 Then Mozaic2 HScroll34
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 11 Then 'wavelines
    .Pic1(PicIdx) = Im
    SaveRedo
    If Check7.Value = 0 Then
        If Option7(0).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 0 'H
        If Option7(1).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 1 'H
        If Option7(2).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 2 'H
    Else
        If Option7(0).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 0 'V
        If Option7(1).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 1 'V
        If Option7(2).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 2 'V
    End If
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll38.Value - 1)
.Pic1(PicIdx).Refresh
End If
'If Cidx = 12 Then 'mix inlaysides with color
'    .Pic3.Picture = Nothing
'    .Pic1(PicIdx) = Im
'    SaveRedo
'        If Option8(0).Value = True Then 'equal background
'            For xx = 0 To .Pic3.Height - 1
'            .Pic3.Line (0, xx)-(.Pic3.Width, xx), Label71.BackColor
'            Next xx
'        End If
        
'    If Option8(1).Value = True Then 'gradient 1
'        Grad .Pic3, Label71.BackColor, Label69.BackColor, Horiz
'    End If
'    If Option8(2).Value = True Then 'gradient 2
'        Grad2 .Pic3, Label71.BackColor, Label69.BackColor, Horiz
'    End If
'    If Option8(3).Value = True Then 'gradient 3
'        Grad3 .Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Horiz
'    End If
'        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll40.Value - 1)
'        AlphaBlend .Pic1(PicIdx).hDC, .Pic3.Width - (SideW + 1), 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll40.Value - 1)
'.Pic1(PicIdx).Refresh
'End If
End With
CDC5.Hide
End Sub

Private Sub Command2_Click() 'cancel
CDC1.Pic1(PicIdx) = Im
CDC5.Hide
End Sub

Private Sub Command3_Click() 'show me
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
If Cidx = 0 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    ChangeColor
End If
If Cidx = 1 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    Changecolor2
End If
If Cidx = 2 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SetLines
End If
If Cidx = 3 Then
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    SetCircles
End If
If Cidx = 4 Then 'contour
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
GetPicData .Pic1(PicIdx)
Contour Label47.BackColor
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 5 Then 'wave x
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    GetPicData .Pic1(PicIdx)
        If Option4(0).Value = True Then
        EffectX .Pic1(PicIdx), HScroll23, HScroll24, 0
        Else
        EffectX .Pic1(PicIdx), HScroll23, HScroll24, 1
        End If
    SetPicData .Pic1(PicIdx)
    .Pic1(PicIdx).Refresh
End If
If Cidx = 6 Then 'wave y
    .Pic3.Picture = Nothing
    .Pic1(PicIdx) = Im
    GetPicData .Pic1(PicIdx)
        If Option5(0).Value = True Then
        EffectY .Pic1(PicIdx), HScroll26, HScroll25, 0
        Else
        EffectY .Pic1(PicIdx), HScroll26, HScroll25, 1
        End If
    SetPicData .Pic1(PicIdx)
    .Pic1(PicIdx).Refresh
End If
If Cidx = 7 Then 'echo
.Pic3.Picture = Nothing
.Pic1(PicIdx) = Im
.Pic3.Picture = .Pic1(PicIdx).Image
.Pic3.Refresh
Echo .Pic1(PicIdx), HScroll27, HScroll28, HScroll29, HScroll30
End If
If Cidx = 8 Then 'tile
.Pic3.Picture = Nothing
.Pic1(PicIdx) = Im
.Pic3.Picture = .Pic1(PicIdx).Image
.Pic3.Refresh
Tile .Pic1(PicIdx), HScroll31, HScroll32
End If
If Cidx = 9 Then 'blinds
    .Pic1(PicIdx) = Im
    GetPicData .Pic1(PicIdx)
    If Check4 = 0 And Option6(0) = True And Check5 = 0 Then 'hor. blinds
    Blinds HScroll33, False, False
    End If
    If Check4 = 1 And Option6(0) = True And Check5 = 0 Then 'hor. blinds reversed
    Blinds HScroll33, True, False
    End If
    If Check4 = 0 And Option6(1) = True And Check5 = 0 Then 'vert. blinds
    Blinds HScroll33, False, True
    End If
    If Check4 = 1 And Option6(1) = True And Check5 = 0 Then 'vert. blinds reversed
    Blinds HScroll33, True, True
    End If
    If Option6(0) = True And Check5 = 1 Then 'hor. bumped blinds
    Blinds3 HScroll33, False
    End If
    If Option6(1) = True And Check5 = 1 Then 'vert. bumped blinds
    Blinds3 HScroll33, True
    End If
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 10 Then 'mosaic
    .Pic1(PicIdx) = Im
    GetPicData .Pic1(PicIdx)
    If Check6 = 0 Then Mozaic HScroll34
    If Check6 = 1 Then Mozaic2 HScroll34
SetPicData .Pic1(PicIdx)
.Pic1(PicIdx).Refresh
End If
If Cidx = 11 Then 'wavelines
    .Pic1(PicIdx) = Im
    If Check7.Value = 0 Then
        If Option7(0).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 0 'H
        If Option7(1).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 1 'H
        If Option7(2).Value = True Then SetWaveLinesH HScroll35, HScroll37, HScroll36, HScroll39, 2 'H
    Else
        If Option7(0).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 0 'V
        If Option7(1).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 1 'V
        If Option7(2).Value = True Then SetWaveLinesV HScroll35, HScroll37, HScroll36, HScroll39, 2 'V
    End If
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll38.Value - 1)
.Pic1(PicIdx).Refresh
End If
'If Cidx = 12 Then 'mix inlaysides with color
'    .Pic3.Picture = Nothing
'    .Pic1(PicIdx) = Im
    
'        If Option8(0).Value = True Then 'equal background
'            For xx = 0 To .Pic3.Height - 1
'            .Pic3.Line (0, xx)-(.Pic3.Width, xx), Label71.BackColor
'            Next xx
'        End If
        
'    If Option8(1).Value = True Then 'gradient 1
'        Grad .Pic3, Label71.BackColor, Label69.BackColor, Horiz
'    End If
'    If Option8(2).Value = True Then 'gradient 2
'        Grad2 .Pic3, Label71.BackColor, Label69.BackColor, Horiz
'    End If
'    If Option8(3).Value = True Then 'gradient 3
'        Grad3 .Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Horiz
'    End If
'        AlphaBlend .Pic1(PicIdx).hDC, 0, 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll40.Value - 1)
'        AlphaBlend .Pic1(PicIdx).hDC, .Pic3.Width - (SideW + 1), 0, SideW, .Pic3.Height, .Pic3.hDC, 0, 0, SideW, .Pic3.Height, Alpha(HScroll40.Value - 1)
'.Pic1(PicIdx).Refresh
'End If
End With
End Sub

Private Sub Command4_Click()
    HScroll20 = CDC1.Pic1(PicIdx).Width / 2
    HScroll21 = CDC1.Pic1(PicIdx).Height / 2
End Sub

Private Sub Command5_Click()
HScroll27 = 5
HScroll28 = 10
HScroll29 = 0
HScroll30 = 0
Check3.Value = 0
End Sub

Private Sub Form_Activate()
Set Im = CDC1.Pic1(PicIdx).Image
On Error Resume Next
For xx = 0 To 12
Frame1(xx).Visible = False
Next xx
Frame1(Cidx).Visible = True
    HScroll3 = CDC1.Pic1(PicIdx).Width / 2
    HScroll4 = CDC1.Pic1(PicIdx).Height / 2
    HScroll19 = 10
    HScroll16 = 1
End Sub

Private Sub Form_Load()
Me.Caption = CDCTitle
Me.Move 10, CDC1.Top, 6180, 3945
Pic1.BackColor = Label3.BackColor
HScroll1 = 5
HScroll2 = 5
HScroll7 = 5
HScroll22 = 5
    Grad Pic2, Label14.BackColor, Label13.BackColor, Vertic
    Grad Pic3, Label71.BackColor, Label69.BackColor, Vertic
    HScroll5 = 250
    HScroll6 = 10
    HScroll8 = 0
    HScroll9 = 0
    HScroll10 = 250
    HScroll15 = 50
    HScroll11 = 250
    HScroll12 = 10
    HScroll13 = 1
    HScroll14 = 25
    HScroll18 = 25
    HScroll17 = 10
    HScroll20 = CDC1.Pic1(PicIdx).Width / 2
    HScroll21 = CDC1.Pic1(PicIdx).Height / 2
    HScroll23 = 50
    HScroll24 = 20
    HScroll26 = 50
    HScroll25 = 20
    Command5_Click
    HScroll31 = 3
    HScroll32 = 3
    HScroll33 = 25
    HScroll34 = 10
    HScroll35 = 25
    HScroll36 = 50
    HScroll37 = 20
    HScroll38 = 5
    HScroll40 = 5
End Sub

Private Sub HScroll1_Change()
Label6 = Format(HScroll1 * 10, "000") & " %"
End Sub

Private Sub HScroll10_Change()
Label30 = Format(HScroll10, "000")
End Sub

Private Sub HScroll11_Change()
Label31 = Format(HScroll11, "000")
End Sub

Private Sub HScroll12_Change()
Label32 = Format(HScroll12, "000")
End Sub

Private Sub HScroll13_Change()
Label33 = HScroll13
End Sub

Private Sub HScroll14_Change()
Label34 = Format(HScroll14, "000")
End Sub

Private Sub HScroll15_Change()
Label35 = Format(HScroll15, "000")
End Sub

Private Sub HScroll16_Change()
Label36 = HScroll16
End Sub

Private Sub HScroll17_Change()
Label37 = Format(HScroll17, "000")
End Sub

Private Sub HScroll18_Change()
Label38 = Format(HScroll18, "000")
End Sub

Private Sub HScroll19_Change()
Label39 = Format(HScroll19 / 10, "0.00")
End Sub

Private Sub HScroll2_Change()
Label11 = Format(HScroll2 * 10, "000") & " %"
End Sub

Private Sub HScroll20_Change()
Label41 = Format(HScroll20, "000")
End Sub

Private Sub HScroll21_Change()
Label40 = Format(HScroll21, "000")
End Sub

Private Sub HScroll22_Change()
Label42 = Format(HScroll22 * 10, "000") & " %"
End Sub

Private Sub HScroll23_Change()
Label49 = Format(HScroll23, "000")
End Sub

Private Sub HScroll24_Change()
Label50 = Format(HScroll24 / 10, "0.0")
End Sub

Private Sub HScroll25_Change()
Label51 = Format(HScroll25 / 10, "0.0")
End Sub

Private Sub HScroll26_Change()
Label52 = Format(HScroll26, "000")
End Sub

Private Sub HScroll27_Change()
Label54 = Format(HScroll27, "00")
End Sub

Private Sub HScroll28_Change()
Label55 = Format(HScroll28, "00") & " %"
End Sub

Private Sub HScroll29_Change()
Label56 = Format(HScroll29, "000")
End Sub

Private Sub HScroll3_Change()
Label19 = Format(HScroll3, "000")
End Sub

Private Sub HScroll30_Change()
Label57 = Format(HScroll30, "000")
End Sub

Private Sub HScroll31_Change()
Label60 = Format(HScroll31, "00")
End Sub

Private Sub HScroll32_Change()
Label61 = Format(HScroll32, "00")
End Sub

Private Sub HScroll33_Change()
Label59 = Format(HScroll33, "000")
End Sub

Private Sub HScroll34_Change()
Label58 = Format(HScroll34, "00")
End Sub

Private Sub HScroll35_Change()
Label62 = Format(HScroll35, "000")
End Sub

Private Sub HScroll36_Change()
Label63 = Format(HScroll36, "000")
End Sub

Private Sub HScroll37_Change()
Label64 = Format(HScroll37 / 10, "0.0")
End Sub

Private Sub HScroll38_Change()
Label67 = Format(HScroll38 * 10, "000") & " %"
End Sub

Private Sub HScroll39_Change()
Label68 = Format(HScroll39, "00")
End Sub

Private Sub HScroll4_Change()
Label20 = Format(HScroll4, "000")
End Sub

Private Sub HScroll40_Change()
Label73 = Format(HScroll40 * 10, "00") & " %"
End Sub

Private Sub HScroll5_Change()
Label21 = Format(HScroll5, "000")
End Sub

Private Sub HScroll6_Change()
Label22 = Format(HScroll6 / 10, "0.00")
End Sub

Private Sub HScroll7_Change()
Label26 = Format(HScroll7 * 10, "000") & " %"
End Sub

Private Sub HScroll8_Change()
Label28 = Format(HScroll8, "000")
End Sub

Private Sub HScroll9_Change()
Label29 = Format(HScroll9, "000")
End Sub

Private Sub Label13_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label13.BackColor
CDC2.CD1.ShowColor
Label13.BackColor = CDC2.CD1.Color
    If Option2(0).Value = True Then
    Grad Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
    If Option2(1).Value = True Then
    Grad2 Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
End Sub

Private Sub Label14_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label14.BackColor
CDC2.CD1.ShowColor
Label14.BackColor = CDC2.CD1.Color
    If Option2(0).Value = True Then
    Grad Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
    If Option2(1).Value = True Then
    Grad2 Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
End Sub

Private Sub Label24_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label24.BackColor
CDC2.CD1.ShowColor
Label24.BackColor = CDC2.CD1.Color
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
    If Option1(3).Value = True Then
    Grad3 Pic1, Label3.BackColor, Label4.BackColor, Label7.BackColor, Vertic
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
    If Option1(3).Value = True Then
    Grad3 Pic1, Label3.BackColor, Label4.BackColor, Label7.BackColor, Vertic
    End If
End Sub

Private Sub Label44_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label44.BackColor
CDC2.CD1.ShowColor
Label44.BackColor = CDC2.CD1.Color
End Sub

Private Sub Label47_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label47.BackColor
CDC2.CD1.ShowColor
Label47.BackColor = CDC2.CD1.Color
End Sub

Private Sub Label66_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label66.BackColor
CDC2.CD1.ShowColor
Label66.BackColor = CDC2.CD1.Color
End Sub

Private Sub Label69_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label69.BackColor
CDC2.CD1.ShowColor
Label69.BackColor = CDC2.CD1.Color
If Option8(0).Value = True Then
Pic3.BackColor = Label71.BackColor
End If
If Option8(1).Value = True Then
Grad Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(2).Value = True Then
Grad2 Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(3).Value = True Then
Grad3 Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Vertic
End If
End Sub

Private Sub Label7_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label7.BackColor
CDC2.CD1.ShowColor
Label7.BackColor = CDC2.CD1.Color
    If Option1(0).Value = True Then
    Pic1.BackColor = Label3.BackColor
    End If
    If Option1(1).Value = True Then
    Grad Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
    If Option1(2).Value = True Then
    Grad2 Pic1, Label3.BackColor, Label4.BackColor, Vertic
    End If
    If Option1(3).Value = True Then
    Grad3 Pic1, Label3.BackColor, Label4.BackColor, Label7.BackColor, Vertic
    End If
End Sub

Private Sub Label71_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label71.BackColor
CDC2.CD1.ShowColor
Label71.BackColor = CDC2.CD1.Color
If Option8(0).Value = True Then
Pic3.BackColor = Label71.BackColor
End If
If Option8(1).Value = True Then
Grad Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(2).Value = True Then
Grad2 Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(3).Value = True Then
Grad3 Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Vertic
End If
End Sub

Private Sub Label74_Click()
CDC2.CD1.CancelError = False
CDC2.CD1.Flags = 3
CDC2.CD1.Color = Label74.BackColor
CDC2.CD1.ShowColor
Label74.BackColor = CDC2.CD1.Color
If Option8(0).Value = True Then
Pic3.BackColor = Label71.BackColor
End If
If Option8(1).Value = True Then
Grad Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(2).Value = True Then
Grad2 Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(3).Value = True Then
Grad3 Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Vertic
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
    If Option1(3).Value = True Then
    Grad3 Pic1, Label3.BackColor, Label4.BackColor, Label7.BackColor, Vertic
    End If
End Sub

Private Sub ChangeColor()
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
If Option1(0).Value = True Then 'equal background
    For xx = 0 To .Pic3.Height - 1
    .Pic3.Line (0, xx)-(.Pic3.Width, xx), Label3.BackColor
    Next xx
End If
If Option1(1).Value = True Then 'gradient 1
    If Check1.Value = 0 Then
    Grad .Pic3, Label3.BackColor, Label4.BackColor, Horiz
    Else
    Grad .Pic3, Label3.BackColor, Label4.BackColor, Vertic
    End If
End If
If Option1(2).Value = True Then 'gradient 2
    If Check1.Value = 0 Then
    Grad2 .Pic3, Label3.BackColor, Label4.BackColor, Horiz
    Else
    Grad2 .Pic3, Label3.BackColor, Label4.BackColor, Vertic
    End If
End If
If Option1(3).Value = True Then 'gradient 3
    If Check1.Value = 0 Then
    Grad3 .Pic3, Label3.BackColor, Label4.BackColor, Label7.BackColor, Horiz
    Else
    Grad3 .Pic3, Label3.BackColor, Label4.BackColor, Label7.BackColor, Vertic
    End If
End If
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll1.Value - 1)
.Pic1(PicIdx).Refresh
End With
End Sub

Private Sub Changecolor2()
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
If Option2(0).Value = True Then 'circ. grad 1
CircleGradient .Pic3, HScroll3, HScroll4, Label14.BackColor, Label13.BackColor, HScroll5, HScroll6 / 10
End If
If Option2(1).Value = True Then 'circ. grad 2
CircleGradient2 .Pic3, HScroll3, HScroll4, Label14.BackColor, Label13.BackColor, HScroll5, HScroll6 / 10
End If
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll2.Value - 1)
.Pic1(PicIdx).Refresh
End With
End Sub
Private Sub Option2_Click(Index As Integer)
    If Option2(0).Value = True Then
    Grad Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
    If Option2(1).Value = True Then
    Grad2 Pic2, Label14.BackColor, Label13.BackColor, Vertic
    End If
End Sub

Private Sub SetLines()
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
CDC1.Pic3.DrawWidth = HScroll13
If Option3(0).Value = True Then
SingleLinesX
End If
If Option3(1).Value = True Then
SingleLinesY
End If
If Option3(2).Value = True Then
SingleLinesX
SingleLinesY
End If
If Option3(3).Value = True Then
SingleLines45
End If
If Option3(4).Value = True Then
SingleLines225
End If
If Option3(5).Value = True Then
SingleLines45
SingleLines225
End If
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll7.Value - 1)
.Pic1(PicIdx).Refresh
.Pic3.DrawWidth = 1
End With
End Sub

Private Sub SingleLinesX()
With CDC1
For xx = 0 To HScroll12 - 1
.Pic3.Line (HScroll8, HScroll9 + (xx * HScroll14))-(HScroll8 + HScroll10, HScroll9 + (xx * HScroll14)), Label24.BackColor
Next xx
End With
End Sub

Private Sub SingleLinesY()
With CDC1
For xx = 0 To HScroll12 - 1
.Pic3.Line (HScroll8 + (xx * HScroll14), HScroll9)-(HScroll8 + (xx * HScroll14), HScroll9 + HScroll11), Label24.BackColor
Next xx
End With
End Sub

Private Sub SingleLines45()
With CDC1
For xx = 0 To HScroll12 - 1
.Pic3.Line (HScroll8 + (xx * HScroll14), HScroll9)-(HScroll8 - HScroll10 + (xx * HScroll14), HScroll9 + HScroll10), Label24.BackColor
Next xx
End With
End Sub

Private Sub SingleLines225()
With CDC1
For xx = 0 To HScroll12 - 1
.Pic3.Line (.Pic3.Width - HScroll8.Value - (xx * HScroll14.Value), HScroll9.Value)-(.Pic3.Width - HScroll8.Value + HScroll11.Value - (xx * HScroll14.Value), HScroll9 + HScroll11), Label24.BackColor
Next xx
End With
End Sub

Private Sub SetCircles()
With CDC1
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
.Pic3.DrawWidth = HScroll16
For xx = 0 To HScroll17 - 1
.Pic3.Circle (HScroll20, HScroll21), HScroll18 + (xx * HScroll15), Label44.BackColor, , , HScroll19 / 10
    If Check2.Value = 1 Then
    .Pic3.Circle (HScroll20, HScroll21), HScroll18 + (xx * HScroll15), Label44.BackColor, , , 1 / (HScroll19 / 10)
    End If

Next xx
AlphaBlend .Pic1(PicIdx).hDC, 0, 0, .Pic3.Width, .Pic3.Height, .Pic3.hDC, 0, 0, .Pic3.Width, .Pic3.Height, Alpha(HScroll22.Value - 1)
.Pic1(PicIdx).Refresh
.Pic3.DrawWidth = 1
End With
End Sub

Private Sub Option8_Click(Index As Integer)
If Option8(0).Value = True Then
Pic3.BackColor = Label71.BackColor
End If
If Option8(1).Value = True Then
Grad Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(2).Value = True Then
Grad2 Pic3, Label71.BackColor, Label69.BackColor, Vertic
End If
If Option8(3).Value = True Then
Grad3 Pic3, Label71.BackColor, Label69.BackColor, Label74.BackColor, Vertic
End If
End Sub
