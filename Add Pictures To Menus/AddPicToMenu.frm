VERSION 5.00
Begin VB.Form AddPicToMenu 
   Caption         =   "Pictures In Menus"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   3120
      Picture         =   "AddPicToMenu.frx":0000
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   3
      Top             =   1080
      Width           =   1740
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   2
      Top             =   3600
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   1320
      Picture         =   "AddPicToMenu.frx":D452
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   1
      Top             =   1080
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Picture To Menus"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mline1 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mline2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "AddPicToMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePatternBrush Lib "GDI32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "User32" () As Long
Private Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuInfo Lib "User32" (ByVal hWnd As Long, mInfo As MENUINFO) As Long
Private Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetMenuInfo Lib "User32" (ByVal hWnd As Long, mInfo As MENUINFO) As Long

Private Type MENUINFO
cbSize              As Long
fMask               As Long
dwStyle             As Long
cyMax               As Long
hbrBack             As Long
dwContextHelpID     As Long
dwMenuData          As Long
End Type

Private Const MIM_BACKGROUND = &H2

Private Sub GradientPictureFill(ByRef Pic As Object, ByVal Col1 As OLE_COLOR, ByVal Col2 As OLE_COLOR, ByVal Fade As Boolean)
Dim I As Long, XY As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim NR1 As Double, NG1 As Double, NB1 As Double
Dim NR2 As Double, NG2 As Double, NB2 As Double

If Fade = True Then XY = Pic.ScaleWidth
If Fade = False Then XY = Pic.ScaleHeight

R1 = Col1 And 255
G1 = ((Col1 \ 256) And 255)
B1 = (Col1 \ 65536) And 255
R2 = Col2 And 255
G2 = ((Col2 \ 256) And 255)
B2 = (Col2 \ 65536) And 255

If Col1 > Col2 Then NR2 = (R1 - R2) / XY: NG2 = (G1 - G2) / XY: NB2 = (B1 - B2) / XY
If Col1 < Col2 Then NR2 = (R2 - R1) / XY: NG2 = (G2 - G1) / XY: NB2 = (B2 - B1) / XY

NR1 = R1: NG1 = G1: NB1 = B1

For I = 0 To XY
If Fade = True Then Pic.Line (I, 0)-(I + 1, Pic.ScaleHeight), RGB(NR1, NG1, NB1), BF
If Fade = False Then Pic.Line (0, I)-(Pic.ScaleWidth, I + 1), RGB(NR1, NG1, NB1), BF
If Col1 > Col2 Then NR1 = NR1 - NR2: NG1 = NG1 - NG2: NB1 = NB1 - NB2
If Col1 < Col2 Then NR1 = NR1 + NR2: NG1 = NG1 + NG2: NB1 = NB1 + NB2
Next
End Sub
 
Private Sub Command1_Click()
Dim MI              As MENUINFO
Dim DC              As Long
Dim Bmp             As Long
Dim OldBmp          As Long
Dim BmpBrush        As Long
Dim MenuBarhWnd     As Long
Dim SubMenuhWnd     As Long
Dim SysMenuhWnd     As Long
Dim DeskTopDC       As Long

'---------------------------------------------------------
MenuBarhWnd = GetMenu(Me.hWnd) 'Get Menu Bar hWnd
SubMenuhWnd = GetSubMenu(MenuBarhWnd, 0) 'Sub Menu "File" hWnd
SysMenuhWnd = GetSystemMenu(Me.hWnd, 0) 'Get Form's System Menu
DeskTopDC = GetWindowDC(GetDesktopWindow()) 'Get Desktop DC
'---------------------------------------------------------
'Paint Picture2 With Gradient Colors
Call GradientPictureFill(Picture2, vbWhite, vbBlue, True)
'---------------------------------------------------------
'Create A Bitamp
DC = CreateCompatibleDC(0) 'Create New DC
Bmp = CreateCompatibleBitmap(DeskTopDC, Picture2.ScaleWidth, Picture2.ScaleHeight) 'Create New Bitmap
OldBmp = SelectObject(DC, Bmp)
'---------------------------------------------------------
'Copy Picture InTo New Bitamp
Call BitBlt(DC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hDC, 0, 0, vbSrcCopy)
'Set MenuInfo
BmpBrush = CreatePatternBrush(Bmp) 'Converts Bitmap to Brush
MI.cbSize = Len(MI) 'Size
MI.fMask = MIM_BACKGROUND 'Flag That Sets/Retrieve Menu Brush
MI.hbrBack = BmpBrush ' Sets New Brush For Menu
'---------------------------------------------------------
'Set New Info To Menu Bar / Or Sub Menus
Call SetMenuInfo(MenuBarhWnd, MI)
'---------------------------------------------------------
Call DrawMenuBar(Me.hWnd) 'RePaints MenuBar
'---------------------------------------------------------
'Delete Bitamp
Bmp = SelectObject(DC, OldBmp)
Call DeleteObject(Bmp)
Call DeleteDC(DC)
'---------------------------------------------------------

'Sub Menu
DC = CreateCompatibleDC(0)
Bmp = CreateCompatibleBitmap(DeskTopDC, Picture1.ScaleWidth, Picture1.ScaleHeight)
OldBmp = SelectObject(DC, Bmp)
Call BitBlt(DC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, vbSrcCopy)
BmpBrush = CreatePatternBrush(Bmp)
MI.cbSize = Len(MI)
MI.fMask = MIM_BACKGROUND
MI.hbrBack = BmpBrush
Call SetMenuInfo(SubMenuhWnd, MI)
Bmp = SelectObject(DC, OldBmp)
Call DeleteObject(Bmp)
Call DeleteDC(DC)

'System Menu
DC = CreateCompatibleDC(0)
Bmp = CreateCompatibleBitmap(DeskTopDC, Picture3.ScaleWidth, Picture3.ScaleHeight)
OldBmp = SelectObject(DC, Bmp)
Call BitBlt(DC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture3.hDC, 0, 0, vbSrcCopy)
BmpBrush = CreatePatternBrush(Bmp)
MI.cbSize = Len(MI)
MI.fMask = MIM_BACKGROUND
MI.hbrBack = BmpBrush
Call SetMenuInfo(SysMenuhWnd, MI)
Bmp = SelectObject(DC, OldBmp)
Call DeleteObject(Bmp)
Call DeleteDC(DC)
End Sub

Private Sub Form_Load()
Picture2.Left = 0
Picture2.Width = Me.ScaleWidth + 100
End Sub
