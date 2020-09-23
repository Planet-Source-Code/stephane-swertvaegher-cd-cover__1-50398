VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CDC1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   524
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3555
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   15
      Top             =   5310
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox TempMem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3150
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   14
      Top             =   5310
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2700
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   5310
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   1980
      TabIndex        =   11
      Top             =   810
      Width           =   2400
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   180
      TabIndex        =   10
      Top             =   1215
      Width           =   1770
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   9
      Top             =   810
      Width           =   1770
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5355
      Top             =   585
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
            Picture         =   "CDC1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CDC1.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CDC1.frx":02B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "kUndo"
            Object.ToolTipText     =   "Undo last action"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kFront"
            Object.ToolTipText     =   "Edit frontside"
            ImageIndex      =   2
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kBack"
            Object.ToolTipText     =   "Edit backside"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7080
      Index           =   1
      Left            =   4590
      ScaleHeight     =   472
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   1
      Top             =   720
      Width           =   7080
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   1230
         Left            =   540
         Top             =   1035
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label SLab1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   495
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label SLab0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   135
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7080
      Index           =   0
      Left            =   4590
      ScaleHeight     =   472
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   0
      Top             =   720
      Width           =   7080
      Begin VB.Label Lab1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   585
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lab0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
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
      Left            =   6075
      TabIndex        =   16
      Top             =   180
      Width           =   3345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Left            =   225
      TabIndex        =   13
      Top             =   5355
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   270
      Stretch         =   -1  'True
      Top             =   5805
      Width           =   1950
   End
   Begin VB.Label Label2 
      Height          =   4470
      Left            =   135
      TabIndex        =   8
      Top             =   720
      Width           =   4290
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   9630
      TabIndex        =   7
      Top             =   180
      Width           =   1365
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New project"
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePic 
         Caption         =   "Save current picture"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenProject 
         Caption         =   "Open project"
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "Save project"
      End
      Begin VB.Menu mnuPrintProject 
         Caption         =   "Print project"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditText 
         Caption         =   "Edit text"
      End
      Begin VB.Menu mnuFullText 
         Caption         =   "Full text"
      End
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Begin VB.Menu mnuClearPic 
         Caption         =   "Clear current picture"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyFit 
         Caption         =   "Copy picture fit"
      End
      Begin VB.Menu mnuCopyAsIs 
         Caption         =   "Copy picture ""as is"""
      End
      Begin VB.Menu mnuCopyScaled 
         Caption         =   "Copy picture scaled"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddPicTube 
         Caption         =   "Add picture tube"
      End
   End
   Begin VB.Menu mnuBG 
      Caption         =   "Background"
      Begin VB.Menu mnuCol1 
         Caption         =   "Normal and linear gradient"
      End
      Begin VB.Menu mnuCol2 
         Caption         =   "Circular gradient"
      End
      Begin VB.Menu mnuAddTexture 
         Caption         =   "Add texture"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuKillR 
         Caption         =   "Kill red component"
      End
      Begin VB.Menu mnuKillG 
         Caption         =   "Kill green component"
      End
      Begin VB.Menu mnuKillB 
         Caption         =   "Kill blue component"
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "Invert colors"
      End
      Begin VB.Menu mnuInvSpec 
         Caption         =   "Invert colors special"
         Begin VB.Menu mnuInvR 
            Caption         =   "Invert red only"
         End
         Begin VB.Menu mnuInvG 
            Caption         =   "Invert green only"
         End
         Begin VB.Menu mnuInvB 
            Caption         =   "Invert blue only"
         End
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwap 
         Caption         =   "Swap colors"
         Begin VB.Menu mnuRBG 
            Caption         =   "RGB --> RBG"
         End
         Begin VB.Menu mnuGRB 
            Caption         =   "RGB --> GRB"
         End
         Begin VB.Menu mnuGBR 
            Caption         =   "RGB --> GBR"
         End
         Begin VB.Menu mnuBRG 
            Caption         =   "RGB --> BRG"
         End
         Begin VB.Menu mnuBGR 
            Caption         =   "RGB --> BGR"
         End
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrey 
         Caption         =   "Greyscale"
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrighten 
         Caption         =   "Brighten picture"
      End
      Begin VB.Menu mnuBrightenMore 
         Caption         =   "Brighten picture more"
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDarken 
         Caption         =   "Darken picture"
      End
      Begin VB.Menu mnuDarkenMore 
         Caption         =   "Darken picture more"
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncrease 
         Caption         =   "Increase colors"
         Begin VB.Menu mnuIncAll 
            Caption         =   "Increase all"
         End
         Begin VB.Menu mnuIncR 
            Caption         =   "Increase red"
         End
         Begin VB.Menu mnuIncG 
            Caption         =   "Increase green"
         End
         Begin VB.Menu mnuIncB 
            Caption         =   "Increase blue"
         End
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDecrease 
         Caption         =   "Decrease colors"
         Begin VB.Menu mnuDecAll 
            Caption         =   "Decrease all"
         End
         Begin VB.Menu mnuDecR 
            Caption         =   "Decrease red"
         End
         Begin VB.Menu mnuDecG 
            Caption         =   "Decrease green"
         End
         Begin VB.Menu mnuDecB 
            Caption         =   "Decrease blue"
         End
      End
      Begin VB.Menu bar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBW 
         Caption         =   "Black and white"
         Begin VB.Menu mnuBW1 
            Caption         =   "Black and white 1"
         End
         Begin VB.Menu mnuBW2 
            Caption         =   "Black and white 2"
         End
         Begin VB.Menu mnuBW3 
            Caption         =   "Black and white 3"
         End
         Begin VB.Menu mnuCharcoal 
            Caption         =   "Charcoal"
         End
         Begin VB.Menu mnuDither 
            Caption         =   "Ordered dither"
         End
         Begin VB.Menu mnuFS1 
            Caption         =   "Floyd Steinberg 1"
         End
         Begin VB.Menu mnuFS2 
            Caption         =   "Floyd Steinberg 2"
         End
         Begin VB.Menu mnuFS3 
            Caption         =   "Floyd Steinberg 3"
         End
         Begin VB.Menu mnuFS4 
            Caption         =   "Floyd Steinberg 4"
         End
         Begin VB.Menu mnuFS5 
            Caption         =   "Floyd Steinberg 5"
         End
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu mnuEmboss 
         Caption         =   "Emboss"
      End
      Begin VB.Menu mnuEmbossSpec 
         Caption         =   "Emboss special"
         Begin VB.Menu mnuHoldR 
            Caption         =   "Hold red"
         End
         Begin VB.Menu mnuHoldG 
            Caption         =   "Hold green"
         End
         Begin VB.Menu mnuHoldB 
            Caption         =   "Hold blue"
         End
      End
      Begin VB.Menu bar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEng 
         Caption         =   "Engrave"
      End
      Begin VB.Menu mnuEngMore 
         Caption         =   "Engrave more"
      End
      Begin VB.Menu bar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiff 
         Caption         =   "Diffuse"
         Begin VB.Menu mnuDiff2 
            Caption         =   "Diffuse X2"
         End
         Begin VB.Menu mnuDiff4 
            Caption         =   "Diffuse X4"
         End
         Begin VB.Menu mnuDiff8 
            Caption         =   "Diffuse X8"
         End
         Begin VB.Menu mnuDiff16 
            Caption         =   "Diffuse X16"
         End
      End
      Begin VB.Menu bar14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelief 
         Caption         =   "Relief"
      End
      Begin VB.Menu bar15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPixelise 
         Caption         =   "Pixelise"
         Begin VB.Menu mnuPix2 
            Caption         =   "Pixelise X2"
         End
         Begin VB.Menu mnuPix4 
            Caption         =   "Pixelise X4"
         End
         Begin VB.Menu mnuPix8 
            Caption         =   "Pixelise X8"
         End
         Begin VB.Menu mnuPix16 
            Caption         =   "Pixelise X16"
         End
      End
      Begin VB.Menu bar16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlur 
         Caption         =   "Blur"
      End
      Begin VB.Menu mnuBlurMore 
         Caption         =   "Blur more"
      End
      Begin VB.Menu bar17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContour 
         Caption         =   "Contour"
      End
      Begin VB.Menu mnuEnEdge 
         Caption         =   "Enchanted edges"
         Begin VB.Menu mnuEdge0 
            Caption         =   "Edge 0"
         End
         Begin VB.Menu mnuEdge1 
            Caption         =   "Edge 1"
         End
         Begin VB.Menu mnuEdge2 
            Caption         =   "Edge 2"
         End
      End
      Begin VB.Menu mnuConnectedContour 
         Caption         =   "Connected contour"
      End
      Begin VB.Menu bar18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoise 
         Caption         =   "Add noise"
      End
      Begin VB.Menu mnuNoiseMore 
         Caption         =   "Add more noise"
      End
      Begin VB.Menu bar19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFog 
         Caption         =   "Add fog"
      End
      Begin VB.Menu mnuFogMore 
         Caption         =   "Add more fog"
      End
      Begin VB.Menu bar20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErode 
         Caption         =   "Erode"
      End
   End
   Begin VB.Menu mnuColFilters 
      Caption         =   "Color filters"
      Begin VB.Menu mnuFreeze 
         Caption         =   "Freeze"
      End
      Begin VB.Menu mnuFreezeMore 
         Caption         =   "Freeze more"
      End
      Begin VB.Menu mnuAfrican 
         Caption         =   "African"
      End
      Begin VB.Menu mnuMoreAf 
         Caption         =   "More African"
      End
      Begin VB.Menu mnuLiquid 
         Caption         =   "Liquid"
      End
      Begin VB.Menu mnuYellow 
         Caption         =   "Yellow"
      End
      Begin VB.Menu mnuDarkMoon 
         Caption         =   "Dark Moon"
      End
      Begin VB.Menu mnuTotEclipse 
         Caption         =   "Total Eclipse"
      End
      Begin VB.Menu mnuPurpleRain 
         Caption         =   "Purple Rain"
      End
      Begin VB.Menu mnuSpooky 
         Caption         =   "Spooky"
      End
      Begin VB.Menu mnuUnreal 
         Caption         =   "Unreal"
      End
      Begin VB.Menu mnuFlame 
         Caption         =   "Flame"
      End
      Begin VB.Menu mnuAquarel 
         Caption         =   "Aquarel"
      End
      Begin VB.Menu mnuSpotted 
         Caption         =   "Spotted"
      End
      Begin VB.Menu mnuRetro 
         Caption         =   "Retro"
      End
      Begin VB.Menu mnuWetPaper 
         Caption         =   "Wet Paper"
      End
      Begin VB.Menu bar21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSoft 
         Caption         =   "Soft colors"
         Begin VB.Menu mnuSoftR 
            Caption         =   "Soft red"
         End
         Begin VB.Menu mnuSoftG 
            Caption         =   "Soft green"
         End
         Begin VB.Menu mnuSoftOrange 
            Caption         =   "Soft orange"
         End
         Begin VB.Menu mnuSoftYellow 
            Caption         =   "Soft yellow"
         End
         Begin VB.Menu mnuSoftPurple 
            Caption         =   "Soft purple"
         End
      End
      Begin VB.Menu bar22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHardCol 
         Caption         =   "Hard colors"
         Begin VB.Menu mnuHardR 
            Caption         =   "Hard red"
         End
         Begin VB.Menu mnuHardG 
            Caption         =   "Hard green"
         End
         Begin VB.Menu mnuHardB 
            Caption         =   "Hard blue"
         End
         Begin VB.Menu mnuHardYellow 
            Caption         =   "Hard yellow"
         End
      End
   End
   Begin VB.Menu mnuBorder 
      Caption         =   "Border"
      Begin VB.Menu mnuAddBorders 
         Caption         =   "Add borders"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effects"
      Begin VB.Menu mnuAddLines 
         Caption         =   "Add lines"
      End
      Begin VB.Menu mnuAddCircles 
         Caption         =   "Add circles"
      End
      Begin VB.Menu mnuAddSinus 
         Caption         =   "Add sinus lines"
      End
      Begin VB.Menu bar23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBlinds 
         Caption         =   "Add blinds"
      End
      Begin VB.Menu bar24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMos 
         Caption         =   "Mosaic"
      End
   End
   Begin VB.Menu mnuDef 
      Caption         =   "Deformation"
      Begin VB.Menu mnuFlipX 
         Caption         =   "Flip X"
      End
      Begin VB.Menu mnuFlipY 
         Caption         =   "Flip Y"
      End
      Begin VB.Menu bar25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMirrorX 
         Caption         =   "Mirror X"
      End
      Begin VB.Menu mnuMirrorXRev 
         Caption         =   "Mirror X reversed"
      End
      Begin VB.Menu mnuMirrorY 
         Caption         =   "Mirror Y"
      End
      Begin VB.Menu mnuMirrorYRev 
         Caption         =   "Mirror Y reversed"
      End
      Begin VB.Menu bar26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWaveX 
         Caption         =   "Wave X"
      End
      Begin VB.Menu mnuWaveY 
         Caption         =   "Wave Y"
      End
      Begin VB.Menu bar27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEcho 
         Caption         =   "Echo"
      End
      Begin VB.Menu bar28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTile 
         Caption         =   "Tile"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "CDC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuEcho_Click()
Cidx = 7
CDC5.Show 1
End Sub

Private Sub mnuFlipX_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, Pic1(PicIdx).ScaleWidth, 0, -Pic3.ScaleWidth, Pic3.ScaleHeight
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuFlipY_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, 0, Pic1(PicIdx).ScaleHeight, Pic3.ScaleWidth, -Pic3.ScaleHeight
Pic1(PicIdx).Refresh
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - (800 * Screen.TwipsPerPixelX)) / 2, 0, 800 * Screen.TwipsPerPixelX, 574 * Screen.TwipsPerPixelY
EnumFonts Printer.hDC, vbNullString, AddressOf EnumFontProc, 0
Screensetup
MakeDirectories
Drive1.Drive = "C:\"
Dir1.Path = Directory & "\CDC_Pictures"
File1.Path = Dir1.Path
File1.Pattern = "*.bmp;*.jpg;*.gif"
CDC1.Show
CDC4.Show
DoEvents
SleepEx 2000, False
CDC4.Hide
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive   ' When drive changes, set directory path.
End Sub
Private Sub Dir1_Change()
   File1.Path = Dir1.Path   ' When directory changes, set file path.
End Sub

Private Sub File1_Click()
Pic2.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
Pic2.Refresh
Dimention Image1, Pic2, Pic2.ScaleWidth, Pic2.ScaleHeight
Label3 = Pic2.Width & " X " & Pic2.Height
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End
End Sub

Private Sub mnuAddBlinds_Click()
Cidx = 9
CDC5.Show 1
End Sub

Private Sub mnuAddBorders_Click()
CDC7.Show 1
End Sub

Private Sub mnuAddCircles_Click()
Cidx = 3
CDC5.Show 1
End Sub

Private Sub mnuAddLines_Click()
Cidx = 2
CDC5.Show 1
End Sub

Private Sub mnuAddPicTube_Click()
BgIdx = 0
CDC8.Show 1
End Sub

Private Sub mnuAddSinus_Click()
Cidx = 11
CDC5.Show 1
End Sub

Private Sub mnuAddTexture_Click()
BgIdx = 1
CDC8.Show 1
End Sub

Private Sub mnuAfrican_Click()
DoFilter 11
End Sub

Private Sub mnuAquarel_Click()
DoFilter 2
End Sub

Private Sub mnuBGR_Click()
DoFilter 4
End Sub

Private Sub mnuBlur_Click()
DoFilter 9
End Sub

Private Sub mnuBlurMore_Click()
DoFilter 10
End Sub

Private Sub mnuBRG_Click()
DoFilter 5
End Sub

Private Sub mnuBrighten_Click()
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
GetPicData Pic1(PicIdx)
BrightenPic 10
SetPicData Pic1(PicIdx)
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuBrightenMore_Click()
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
GetPicData Pic1(PicIdx)
BrightenPic 50
SetPicData Pic1(PicIdx)
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuBW1_Click()
DoFilter 13
End Sub

Private Sub mnuBW2_Click()
DoFilter 14
End Sub

Private Sub mnuBW3_Click()
DoFilter 15
End Sub

Private Sub mnuCharcoal_Click()
DoFilter 16
End Sub

Private Sub mnuClearPic_Click()
Temp = MsgBox("The whole picture will be cleared. Continue ?", vbQuestion + vbYesNo, BLMTitle)
If Temp = vbNo Then Exit Sub
SaveRedo
Pic1(PicIdx).Picture = Nothing
End Sub

Private Sub mnuCol1_Click()
Cidx = 0
CDC5.Show 1
End Sub

Private Sub mnuCol2_Click()
Cidx = 1
CDC5.Show 1
End Sub

Private Sub mnuConnectedContour_Click()
DoFilter 58
End Sub

Private Sub mnuContour_Click()
If PicPresent(CDC1) = False Then Exit Sub
Cidx = 4
CDC5.Show 1
End Sub

Private Sub mnuCopyAsIs_Click()
CopyToMap Pic1(PicIdx), Pic2.ScaleWidth, Pic2.ScaleHeight, (PicW - Pic2.ScaleWidth) / 2, (PicH - Pic2.ScaleHeight) / 2
End Sub

Private Sub mnuCopyFit_Click()
CopyToMap Pic1(PicIdx), PicW, PicH, 0, 0
End Sub

Private Sub mnuCopyScaled_Click()
If PicPresent(CDC1) = False Then Exit Sub
CDC3.Show 1
End Sub

Private Sub mnuDarken_Click()
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
GetPicData Pic1(PicIdx)
BrightenPic -10
SetPicData Pic1(PicIdx)
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuDarkenMore_Click()
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
GetPicData Pic1(PicIdx)
BrightenPic -50
SetPicData Pic1(PicIdx)
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuDarkMoon_Click()
DoFilter 61
End Sub

Private Sub mnuDecAll_Click()
DoFilter 81
End Sub

Private Sub mnuDecB_Click() 'decrease blue
DoFilter 23
End Sub

Private Sub mnuDecG_Click() 'decrease green
DoFilter 24
End Sub

Private Sub mnuDecR_Click() 'decrease red
DoFilter 25
End Sub

Private Sub mnuDiff16_Click()
DoFilter 46
End Sub

Private Sub mnuDiff2_Click()
DoFilter 43
End Sub

Private Sub mnuDiff4_Click()
DoFilter 44
End Sub

Private Sub mnuDiff8_Click()
DoFilter 45
End Sub

Private Sub mnuDither_Click()
DoFilter 22
End Sub

Private Sub mnuEdge0_Click()
DoFilter 55
End Sub

Private Sub mnuEdge1_Click()
DoFilter 56
End Sub

Private Sub mnuEdge2_Click()
DoFilter 57
End Sub

Private Sub mnuEditText_Click()
CDC2.Show 1
End Sub

Private Sub mnuEmboss_Click()
DoFilter 37
End Sub

Private Sub mnuEng_Click()
DoFilter 41
End Sub

Private Sub mnuEngMore_Click()
DoFilter 42
End Sub

Private Sub mnuErode_Click()
DoFilter 54
End Sub

Private Sub mnuFlame_Click()
DoFilter 68
End Sub

Private Sub mnuFog_Click()
DoFilter 52
End Sub

Private Sub mnuFogMore_Click()
DoFilter 53
End Sub

Private Sub mnuFreeze_Click()
DoFilter 59
End Sub

Private Sub mnuFreezeMore_Click()
DoFilter 60
End Sub

Private Sub mnuFS1_Click()
DoFilter 17
End Sub

Private Sub mnuFS2_Click()
DoFilter 18
End Sub

Private Sub mnuFS3_Click()
DoFilter 19
End Sub

Private Sub mnuFS4_Click()
DoFilter 20
End Sub

Private Sub mnuFS5_Click()
DoFilter 21
End Sub

Private Sub mnuFullText_Click()
CDC9.Show 1
End Sub

Private Sub mnuGBR_Click()
DoFilter 6
End Sub

Private Sub mnuGRB_Click()
DoFilter 7
End Sub

Private Sub mnuGrey_Click()
DoFilter 36
End Sub

Private Sub mnuHardB_Click()
DoFilter 78
End Sub

Private Sub mnuHardG_Click()
DoFilter 77
End Sub

Private Sub mnuHardR_Click()
DoFilter 76
End Sub

Private Sub mnuHardYellow_Click()
DoFilter 79
End Sub

Private Sub mnuHoldB_Click() 'hold blue
DoFilter 40
End Sub

Private Sub mnuHoldG_Click() 'hold green
DoFilter 39
End Sub

Private Sub mnuHoldR_Click() 'hold red
DoFilter 38
End Sub

Private Sub mnuIncAll_Click() 'increase all
DoFilter 80
End Sub

Private Sub mnuIncB_Click() 'increase blue
DoFilter 26
End Sub

Private Sub mnuIncG_Click() 'increase green
DoFilter 27
End Sub

Private Sub mnuIncR_Click() 'increase red
DoFilter 28
End Sub

Private Sub mnuInfo_Click()
CDC4.Height = 334 * Screen.TwipsPerPixelY
CDC4.Show 1
End Sub

Private Sub mnuInvB_Click()
DoFilter 33
End Sub

Private Sub mnuInvert_Click()
DoFilter 32
End Sub

Private Sub mnuInvG_Click()
DoFilter 34
End Sub

Private Sub mnuInvR_Click()
DoFilter 35
End Sub

Private Sub mnuKillB_Click()
DoFilter 29
End Sub

Private Sub mnuKillG_Click()
DoFilter 30
End Sub

Private Sub mnuKillR_Click()
DoFilter 31
End Sub

Private Sub mnuLiquid_Click()
DoFilter 3
End Sub

Private Sub mnuMirrorX_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, Pic1(PicIdx).ScaleWidth, 0, -Pic3.ScaleWidth, Pic3.ScaleHeight
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, Pic1(PicIdx).ScaleWidth / 2, 0, -Pic3.ScaleWidth / 2, Pic3.ScaleHeight, Pic1(PicIdx).ScaleWidth / 2, 0
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuMirrorXRev_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, Pic1(PicIdx).ScaleWidth / 2, 0, -Pic3.ScaleWidth / 2, Pic3.ScaleHeight, Pic1(PicIdx).ScaleWidth / 2, 0
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuMirrorY_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, 0, Pic1(PicIdx).ScaleHeight, Pic3.ScaleWidth, -Pic3.ScaleHeight
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, 0, Pic1(PicIdx).ScaleHeight / 2, Pic3.ScaleWidth, -Pic3.ScaleHeight / 2, 0, Pic1(PicIdx).ScaleHeight / 2
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuMirrorYRev_Click()
SaveRedo
Pic3.Width = Pic1(PicIdx).Width
Pic3.Height = Pic1(PicIdx).Height
Pic3.Picture = Pic1(PicIdx).Image
Pic1(PicIdx).PaintPicture Pic3.Picture, 0, Pic1(PicIdx).ScaleHeight / 2, Pic3.ScaleWidth, -Pic3.ScaleHeight / 2, 0, Pic1(PicIdx).ScaleHeight / 2
Pic1(PicIdx).Refresh
End Sub

Private Sub mnuMoreAf_Click()
DoFilter 12
End Sub

Private Sub mnuMos_Click()
Cidx = 10
CDC5.Show 1
End Sub

Private Sub mnuNew_Click()
Temp = MsgBox("Do you want to start a new project ?" & vbCrLf & vbCrLf & "All pictures and text will be cleared." & vbCrLf & vbCrLf & "Continue ?", vbYesNo + vbQuestion + vbDefaultButton2, CDCTitle)
If Temp = vbNo Then Exit Sub
' clear and set all text
DefaultTextPositions
Pic1(0).Picture = LoadPicture("")
Pic1(1).Picture = LoadPicture("")
Toolbar1.Buttons(3).Value = tbrPressed
Label1 = "FRONTSIDE"
Pic1(0).Visible = True
Pic1(1).Visible = False
PicIdx = 0
Re(0) = False
Re(1) = False
Toolbar1.Buttons(1).Enabled = False
End Sub

Private Sub File1_DblClick()
CopyToMap Pic1(PicIdx), PicW, PicH, 0, 0
End Sub

Private Sub mnuNoisemore_Click()
DoFilter 1
End Sub

Private Sub mnuNoise_Click()
DoFilter 0
End Sub

Private Sub mnuOpenProject_Click()
Temp = MsgBox("Open a new project" & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, CDCTitle)
If Temp = vbNo Then Exit Sub
OpenProject

End Sub

Private Sub mnuPix16_Click()
DoFilter 51
End Sub

Private Sub mnuPix2_Click()
DoFilter 48
End Sub

Private Sub mnuPix4_Click()
DoFilter 49
End Sub

Private Sub mnuPix8_Click()
DoFilter 50
End Sub

Private Sub mnuPrintProject_Click()
Temp = MsgBox("This will print the front and backside" & vbCrLf & "of the CD-cover" & vbCrLf & vbCrLf & "Do you wish to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, CDCTitle)
If Temp = vbNo Then Exit Sub
PrtposX(0) = 450
PrtposY(0) = 450
PrintFront
End Sub

Private Sub mnuPurpleRain_Click()
DoFilter 65
End Sub

Private Sub mnuRBG_Click()
DoFilter 8
End Sub

Private Sub mnuRelief_Click()
DoFilter 47
End Sub

Private Sub mnuRetro_Click()
DoFilter 69
End Sub

Private Sub mnuSavePic_Click()
On Error GoTo SavePic
CDC2.CD1.CancelError = True
CDC2.CD1.Flags = 2
CDC2.CD1.InitDir = App.Path & "\CDC-SavedPictures"
CDC2.CD1.DialogTitle = CDCTitle & " - Save picture"
CDC2.CD1.Filter = "Picture|*.bmp"
CDC2.CD1.FileName = ""
CDC2.CD1.ShowSave
If CDC2.CD1.FileName = "" Then Exit Sub
    SavePicture CDC1.Pic1(PicIdx).Image, CDC2.CD1.FileName
    File1.Refresh
SavePic:
End Sub

Private Sub mnuSaveProject_Click()
SaveProject
End Sub

Private Sub mnuSoftG_Click()
DoFilter 72
End Sub

Private Sub mnuSoftOrange_Click()
DoFilter 73
End Sub

Private Sub mnuSoftPurple_Click()
DoFilter 75
End Sub

Private Sub mnuSoftR_Click()
DoFilter 71
End Sub

Private Sub mnuSoftYellow_Click()
DoFilter 74
End Sub

Private Sub mnuSpooky_Click()
DoFilter 66
End Sub

Private Sub mnuSpotted_Click()
DoFilter 67
End Sub

Private Sub mnuTile_Click()
Cidx = 8
CDC5.Show 1
End Sub

Private Sub mnuTotEclipse_Click()
DoFilter 63
End Sub

Private Sub mnuUnreal_Click()
DoFilter 64
End Sub

Private Sub mnuWaveX_Click()
Cidx = 5
CDC5.Show 1
End Sub

Private Sub mnuWaveY_Click()
Cidx = 6
CDC5.Show 1
End Sub

Private Sub mnuWetPaper_Click()
DoFilter 70
End Sub

Private Sub mnuYellow_Click()
DoFilter 62
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "kUndo"
Redo
Case "kFront"
Label1 = "FRONTSIDE"
Pic1(0).Visible = True
Pic1(1).Visible = False
PicIdx = 0
Case "kBack"
Label1 = "BACKSIDE"
Pic1(1).Visible = True
Pic1(0).Visible = False
PicIdx = 1
End Select
End Sub

Private Sub DoFilter(Fi%)
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
GetPicData Pic1(PicIdx)
Select Case Fi
Case 0 'noise
AddNoise 50
Case 1 'more noise
AddNoise 100
Case 2 'aquarel
Aquarel
Case 3 'liquid
Liquid
Case 4
BGR
Case 5
BRG
Case 6
GBR
Case 7
GRB
Case 8
RBG
Case 9
Blur
Case 10
BlurMore
Case 11
Brown 128
Case 12
Brown 256
Case 13
NearestColorBW &H303030
Case 14
NearestColorBW &H808080
Case 15
NearestColorBW &HC0C0C0
Case 16
Charcoal
Case 17
FloydSteinberg 3
Case 18
FloydSteinberg 6
Case 19
FloydSteinberg 10
Case 20
FloydSteinberg 15
Case 21
FloydSteinberg 20
Case 22
OrderedDither
Case 23
Decrease 3
Case 24
Decrease 2
Case 25
Decrease 1
Case 26
Increase 3
Case 27
Increase 2
Case 28
Increase 1
Case 29
KillComp 1
Case 30
KillComp 2
Case 31
KillComp 3
Case 32
NegativeImage 0
Case 33
NegativeImage 3
Case 34
NegativeImage 2
Case 35
NegativeImage 1
Case 36
GreyScale
Case 37
Emboss 128, 128, 128
Case 38
EmbossHR
Case 39
EmbossHG
Case 40
EmbossHB
Case 41
Engrave
Case 42
EngraveMore
Case 43
Diffuse 2
Case 44
Diffuse 4
Case 45
Diffuse 8
Case 46
Diffuse 16
Case 47
Relief
Case 48
Pixelize 2
Case 49
Pixelize 4
Case 50
Pixelize 8
Case 51
Pixelize 16
Case 52
Fog 50
Case 53
Fog 120
Case 54
Erode 30
Case 55
EdgeEnhance 0
Case 56
EdgeEnhance 1
Case 57
EdgeEnhance 2
Case 58
ConnectedContour
Case 59
Freeze 1.1
Case 60
Freeze 1.5
Case 61
DarkMoon
Case 62
Yellow
Case 63
TotalEclipse
Case 64
UnReal
Case 65
PurpleRain
Case 66
Spooky
Case 67
Effect0 6
Case 68
Flame
Case 69
Effect0 7
Case 70
Effect0 11
Case 71
Effect0 3
Case 72
Effect0 1
Case 73
Effect0 10
Case 74
Effect0 9
Case 75
Effect0 2
Case 76
Effect0 4
Case 77
Effect0 5
Case 78
Effect0 0
Case 79
Effect0 8
Case 80 'increase all
Increase 4
Case 81 ' decrease all
Decrease 4
End Select
SetPicData Pic1(PicIdx)
Pic1(PicIdx).Refresh
End Sub

Private Sub SaveProject()
With CDC1
On Error GoTo SaveProject0
CDC2.CD1.CancelError = True
CDC2.CD1.Flags = 2
CDC2.CD1.InitDir = App.Path & "\CDC-Projects"
CDC2.CD1.DialogTitle = CDCTitle & " - Save project"
CDC2.CD1.Filter = "Booklet|*.cdb|All files|*.*"
CDC2.CD1.FileName = ProjectName
CDC2.CD1.ShowSave
    ff = FreeFile
    Open CDC2.CD1.FileName For Output As #ff
    For xx = 0 To 49
            Print #ff, .Lab0(xx).Caption
            Print #ff, .Lab1(xx).Caption
            
            Print #ff, .Lab0(xx).FontName
            Print #ff, .Lab1(xx).FontName
            
            Print #ff, Int(.Lab0(xx).FontSize)
            Print #ff, Int(.Lab1(xx).FontSize)
            
            If .Lab0(xx).FontBold = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            If .Lab1(xx).FontBold = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            
            If .Lab0(xx).FontItalic = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            If .Lab1(xx).FontItalic = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            
            If .Lab0(xx).FontUnderline = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            If .Lab1(xx).FontUnderline = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            
            Print #ff, .Lab0(xx).ForeColor
            Print #ff, .Lab1(xx).ForeColor

            If .Lab0(xx).Visible = True Then
            Print #ff, 1
            Else
            Print #ff, 0
            End If
            If .Lab1(xx).Visible = True Then
            Print #ff, 1
            Else
            Print #ff, 0
            End If

            Print #ff, Int(.Lab0(xx).Left)
            Print #ff, Int(.Lab1(xx).Left)
            Print #ff, Int(.Lab0(xx).Top)
            Print #ff, Int(.Lab1(xx).Top)

            If .SLab0(xx).Visible = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            If .SLab1(xx).Visible = False Then
            Print #ff, 0
            Else
            Print #ff, 1
            End If
            
            Print #ff, .SLab0(xx).ForeColor
            Print #ff, .SLab1(xx).ForeColor

            Print #ff, ShX(0, xx)
            Print #ff, ShX(1, xx)
            Print #ff, ShY(0, xx)
            Print #ff, ShY(1, xx)
    Next xx
Close #ff
    PrTitle = Left(CDC2.CD1.FileName, Len(CDC2.CD1.FileName) - 4)
    SavePicture .Pic1(0).Image, PrTitle & "1.bmp"
    SavePicture .Pic1(1).Image, PrTitle & "2.bmp"
ProjectName = CDC2.CD1.FileTitle
.Label4 = ProjectName
End With
Exit Sub
SaveProject0:
Close #ff
End Sub

Private Sub OpenProject()
With CDC1
On Error GoTo OpenProject0
CDC2.CD1.CancelError = True
CDC2.CD1.Flags = 2
CDC2.CD1.InitDir = App.Path & "\CDC-Projects"
CDC2.CD1.DialogTitle = CDCTitle & " - Open project"
CDC2.CD1.Filter = "Booklet|*.cdb|All files|*.*"
CDC2.CD1.FileName = ""
CDC2.CD1.ShowOpen
    ff = FreeFile
    Open CDC2.CD1.FileName For Input As #ff
    For xx = 0 To 49
            'load text
            Line Input #ff, Temp
            .Lab0(xx).Caption = Temp
            .SLab0(xx).Caption = Temp
            Line Input #ff, Temp
            .Lab1(xx).Caption = Temp
            .SLab1(xx).Caption = Temp
            'load fontnames
            Line Input #ff, Temp
            .Lab0(xx).FontName = Temp
            .SLab0(xx).FontName = Temp
            Line Input #ff, Temp
            .Lab1(xx).FontName = Temp
            .SLab1(xx).FontName = Temp
            'load fontsizes
            Input #ff, yy
            .Lab0(xx).FontSize = yy
            .SLab0(xx).FontSize = yy
            Input #ff, yy
            .Lab1(xx).FontSize = yy
            .SLab1(xx).FontSize = yy
            'load fontbold
            Input #ff, yy
                If yy = 0 Then
                .Lab0(xx).FontBold = False
                Else
                .Lab0(xx).FontBold = True
                End If
                .SLab0(xx).FontBold = .Lab0(xx).FontBold
            Input #ff, yy
                If yy = 0 Then
                .Lab1(xx).FontBold = False
                Else
                .Lab1(xx).FontBold = True
                End If
                .SLab1(xx).FontBold = .Lab1(xx).FontBold
             'load fontitalic
            Input #ff, yy
                If yy = 0 Then
                .Lab0(xx).FontItalic = False
                Else
                .Lab0(xx).FontItalic = True
                End If
                .SLab0(xx).FontItalic = .Lab0(xx).FontItalic
            Input #ff, yy
                If yy = 0 Then
                .Lab1(xx).FontItalic = False
                Else
                .Lab1(xx).FontItalic = True
                End If
                .SLab1(xx).FontItalic = .Lab1(xx).FontItalic
            'load fontunderline
            Input #ff, yy
                If yy = 0 Then
                .Lab0(xx).FontUnderline = False
                Else
                .Lab0(xx).FontUnderline = True
                End If
                .SLab0(xx).FontUnderline = .Lab0(xx).FontUnderline
            Input #ff, yy
                If yy = 0 Then
                .Lab1(xx).FontUnderline = False
                Else
                .Lab1(xx).FontUnderline = True
                End If
                .SLab1(xx).FontUnderline = .Lab1(xx).FontUnderline
            'load colors text
            Input #ff, Num
            .Lab0(xx).ForeColor = Num
            Input #ff, Num
            .Lab1(xx).ForeColor = Num
            'load visible text
            Input #ff, yy
                If yy = 0 Then
                Lab0(xx).Visible = False
                Else
                Lab0(xx).Visible = True
                End If
            Input #ff, yy
                If yy = 0 Then
                Lab1(xx).Visible = False
                Else
                Lab1(xx).Visible = True
                End If
           'load left text
           Input #ff, yy
           .Lab0(xx).Left = yy
           Input #ff, yy
           .Lab1(xx).Left = yy
           'load top text
           Input #ff, yy
           .Lab0(xx).Top = yy
           Input #ff, yy
           .Lab1(xx).Top = yy
           'load visible shadow
            Input #ff, yy
                If yy = 0 Then
                SLab0(xx).Visible = False
                Else
                SLab0(xx).Visible = True
                End If
            Input #ff, yy
                If yy = 0 Then
                SLab1(xx).Visible = False
                Else
                SLab1(xx).Visible = True
                End If
           'load colors shadow
            Input #ff, Num
            .SLab0(xx).ForeColor = Num
            Input #ff, Num
            .SLab1(xx).ForeColor = Num
            'load shadowpositions
            Input #ff, ShX(0, xx)
            Input #ff, ShX(1, xx)
            Input #ff, ShY(0, xx)
            Input #ff, ShY(1, xx)
            .SLab0(xx).Left = .Lab0(xx).Left + ShX(0, xx)
            .SLab1(xx).Left = .Lab1(xx).Left + ShX(1, xx)
            .SLab0(xx).Top = .Lab0(xx).Top + ShY(0, xx)
            .SLab1(xx).Top = .Lab1(xx).Top + ShY(1, xx)
    Next xx
Close #ff
'load pictures
    PrTitle = Left(CDC2.CD1.FileName, Len(CDC2.CD1.FileName) - 4)
    .Pic1(0).Picture = LoadPicture(PrTitle & "1.bmp")
    .Pic1(1).Picture = LoadPicture(PrTitle & "2.bmp")
'load projectname
ProjectName = CDC2.CD1.FileTitle
.Label4 = ProjectName
'do screensettings
.Toolbar1.Buttons(3).Value = tbrPressed
Pic1(0).Visible = True
Pic1(1).Visible = False
Label1 = "FRONTSIDE"
PicIdx = 0
Re(0) = False
Re(1) = False
Toolbar1.Buttons(1).Enabled = False
End With
'cdc9.Text1 = ""
OpenProject0:
Close #ff
End Sub

