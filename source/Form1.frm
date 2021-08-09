VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Start Menu"
   ClientHeight    =   10020
   ClientLeft      =   7005
   ClientTop       =   2880
   ClientWidth     =   11925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox File_Icon6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   27
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox File_Icon5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   26
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox File_Icon4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   25
      Top             =   7080
      Width           =   495
   End
   Begin VB.PictureBox File_Icon3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   24
      Top             =   7080
      Width           =   495
   End
   Begin VB.PictureBox File_Icon2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox File_Icon1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   22
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox App_Icon12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon7 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon8 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3840
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "Search"
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   2400
      Width           =   6855
      Begin VB.TextBox Search_Input 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Type Here To Search"
         Top             =   120
         Width           =   6015
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   0
         Picture         =   "Form1.frx":4DD03
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6735
      End
      Begin VB.Line Search_UnderLine 
         BorderColor     =   &H8000000D&
         X1              =   0
         X2              =   6720
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Label File_Name6 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6720
      TabIndex        =   39
      Top             =   7920
      Width           =   2370
   End
   Begin VB.Label File_Date6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   6720
      TabIndex        =   38
      Top             =   8160
      Width           =   870
   End
   Begin VB.Label File_Name5 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   37
      Top             =   7920
      Width           =   2370
   End
   Begin VB.Label File_Date5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3480
      TabIndex        =   36
      Top             =   8160
      Width           =   870
   End
   Begin VB.Label File_Name4 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6720
      TabIndex        =   35
      Top             =   7080
      Width           =   2370
   End
   Begin VB.Label File_Date4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   6720
      TabIndex        =   34
      Top             =   7320
      Width           =   870
   End
   Begin VB.Label File_Name3 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   33
      Top             =   7080
      Width           =   2370
   End
   Begin VB.Label File_Date3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3480
      TabIndex        =   32
      Top             =   7320
      Width           =   870
   End
   Begin VB.Label File_Name2 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6720
      TabIndex        =   31
      Top             =   6240
      Width           =   2370
   End
   Begin VB.Label File_Date2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   6720
      TabIndex        =   30
      Top             =   6480
      Width           =   870
   End
   Begin VB.Label File_Date1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Date"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3480
      TabIndex        =   29
      Top             =   6480
      Width           =   870
   End
   Begin VB.Label File_Name1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Title"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   28
      Top             =   6240
      Width           =   2370
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000B&
      X1              =   9360
      X2              =   8280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label More_Recommended_Btn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More  >"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   21
      Top             =   5760
      Width           =   945
   End
   Begin VB.Image More_Recommended_Btn_Bg 
      Height          =   330
      Left            =   7680
      Picture         =   "Form1.frx":8184D
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label All_Apps_Btn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Apps >"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   20
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Image All_Apps_Btn_Bg 
      Height          =   330
      Left            =   7440
      Picture         =   "Form1.frx":98797
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label Type_Title2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recommended"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   19
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Close_Mask_Btn_Right 
      BackStyle       =   0  'Transparent
      Height          =   9855
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Close_Mask_Btn_Top 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   7920
      Picture         =   "Form1.frx":AF6E1
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label User_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   9240
      Width           =   555
   End
   Begin VB.Image User_Image 
      Height          =   600
      Left            =   2760
      Picture         =   "Form1.frx":DAA43
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   570
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   2040
      X2              =   9480
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Type_Title1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pinned"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000B&
      X1              =   2040
      X2              =   2040
      Y1              =   8160
      Y2              =   9720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   2160
      X2              =   9360
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   9480
      X2              =   9480
      Y1              =   2280
      Y2              =   8760
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Close_Mask_Btn_Left 
      BackStyle       =   0  'Transparent
      Height          =   9855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'透明窗口
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'窗口置顶
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'读取ini文件
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
'调用快捷方式
Private Declare Function SHGetFileInfo Lib _
   "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
'获取文件图标
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const ILD_TRANSPARENT = &H1
Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 260
   szTypeName As String * 80
End Type
Private udtFI As SHFILEINFO
'获取文件图标
 
Public Sub ShowIcon(strAnyFile As String, pic As PictureBox)
    Dim hLarge As Long
    pic.AutoRedraw = True
    pic.Cls
    hLarge = SHGetFileInfo(strAnyFile, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
    ImageList_Draw hLarge, udtFI.iIcon, pic.hDC, 0, 0, ILD_TRANSPARENT
    pic.Refresh
End Sub
'获取文件图标

Public Sub Start(appid As String)
    Dim read_OK As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("Pinned", appid, "Noting", read2, 256, App.Path & "\AppList.ini")
    a = ShellExecute(Me.hwnd, vbNullString, App.Path & "\Links\" & read2, vbNullString, "D:\", SW_SHOWNORMAL)
    End
End Sub
'打开指定文件快捷方式

Public Sub PlaceIcon(appid As String, pic As PictureBox)
    Dim read_OK As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("Pinned", appid, "Noting", read2, 256, App.Path & "\AppList.ini")
    Call ShowIcon(App.Path & "\Links\" & read2, pic)
    pic.ToolTipText = read2
End Sub
'添加Icon

Public Sub Get_Latest_File(file_name As Label, file_date As Label, file_icon As PictureBox)
    Dim d As String, t As Date, s As String, ld As String
    'Dim read_OK As Long
    'Dim read2 As String
    'read2 = String(255, 0)
    'read_OK = GetPrivateProfileString("Recommended", "ShareFolder", "Noting", read2, 256, App.Path & "\AppList.ini")
    read2 = App.Path
    d = Dir(read2 & "\*.*")
    Do Until d = ""
        If FileDateTime(read2 & "\" & d) > t Then
            t = FileDateTime(read2 & "\" & d)
            s = read2 & "\" & d
            ld = d
        End If
        d = Dir
    Loop
    file_name.Caption = ld
    file_date.Caption = FileDateTime(s)
    Call ShowIcon(s, file_icon)
End Sub

Private Sub Close_Mask_Btn_Left_Click()
    End
End Sub

Private Sub Close_Mask_Btn_Right_Click()
    End
End Sub

Private Sub Close_Mask_Btn_Top_Click()
    End
End Sub

Private Sub Form_Load()
    Me.BackColor = &HFF0000
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HFF0000, 0, LWA_COLORKEY
    
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1
    
    Form1.Left = (Screen.Width - Form1.Width) / 2
    Form1.Top = Screen.Height - 550 - Form1.Height
    
    Call PlaceIcon("App1", App_Icon1)
    Call PlaceIcon("App2", App_Icon2)
    Call PlaceIcon("App3", App_Icon3)
    Call PlaceIcon("App4", App_Icon4)
    Call PlaceIcon("App5", App_Icon5)
    Call PlaceIcon("App6", App_Icon6)
    Call PlaceIcon("App7", App_Icon7)
    Call PlaceIcon("App8", App_Icon8)
    Call PlaceIcon("App9", App_Icon9)
    Call PlaceIcon("App10", App_Icon10)
    Call PlaceIcon("App11", App_Icon11)
    Call PlaceIcon("App12", App_Icon12)
    
    Call Get_Latest_File(File_Name1, File_Date1, File_Icon1)
    Call Get_Latest_File(File_Name2, File_Date2, File_Icon2)
    Call Get_Latest_File(File_Name3, File_Date3, File_Icon3)
    Call Get_Latest_File(File_Name4, File_Date4, File_Icon4)
    Call Get_Latest_File(File_Name5, File_Date5, File_Icon5)
    Call Get_Latest_File(File_Name6, File_Date6, File_Icon6)
End Sub

Private Sub App_Icon1_Click()
    Start ("App1")
End Sub

Private Sub App_Icon2_Click()
    Start ("App2")
End Sub

Private Sub App_Icon3_Click()
    Start ("App3")
End Sub

Private Sub App_Icon4_Click()
    Start ("App4")
End Sub

Private Sub App_Icon5_Click()
    Start ("App5")
End Sub

Private Sub App_Icon6_Click()
    Start ("App6")
End Sub

Private Sub App_Icon7_Click()
    Start ("App7")
End Sub

Private Sub App_Icon8_Click()
    Start ("App8")
End Sub

Private Sub App_Icon9_Click()
    Start ("App9")
End Sub

Private Sub App_Icon10_Click()
    Start ("App10")
End Sub

Private Sub App_Icon11_Click()
    Start ("App11")
End Sub

Private Sub App_Icon12_Click()
    Start ("App12")
End Sub

