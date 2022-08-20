VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Windows Start Menu"
   ClientHeight    =   10020
   ClientLeft      =   7005
   ClientTop       =   2880
   ClientWidth     =   11895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10000
   ScaleMode       =   0  'User
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.Timer AnimationTimer 
      Interval        =   1
      Left            =   5520
      Top             =   9360
   End
   Begin VB.TextBox Search_Input 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      ToolTipText     =   "Type Here To Search"
      Top             =   2400
      Width           =   6015
   End
   Begin VB.PictureBox File_Icon6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   25
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox File_Icon5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   24
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox File_Icon4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   7080
      Width           =   495
   End
   Begin VB.PictureBox File_Icon3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   22
      Top             =   7080
      Width           =   495
   End
   Begin VB.PictureBox File_Icon2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   21
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox File_Icon1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   20
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox App_Icon12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.PictureBox App_Icon6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8160
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   3840
      Width           =   495
   End
   Begin VB.PictureBox App_Icon1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label PowerPtn 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   7920
      TabIndex        =   27
      Top             =   9240
      Width           =   615
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
      TabIndex        =   19
      Top             =   5760
      Width           =   945
   End
   Begin VB.Image More_Recommended_Btn_Bg 
      Height          =   330
      Left            =   7680
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label All_Apps_Btn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7560
      TabIndex        =   18
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Image All_Apps_Btn_Bg 
      Height          =   330
      Left            =   7440
      Picture         =   "Form1.frx":16F56
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
      TabIndex        =   17
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Close_Mask_Btn_Right 
      BackStyle       =   0  'Transparent
      Height          =   9855
      Left            =   9600
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Close_Mask_Btn_Top 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label UserInfoBtn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   1
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   1920
      X2              =   9480
      Y1              =   9101.797
      Y2              =   9101.797
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
      TabIndex        =   0
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Close_Mask_Btn_Left 
      BackStyle       =   0  'Transparent
      Height          =   9855
      Left            =   120
      TabIndex        =   2
      Top             =   -1320
      Width           =   1695
   End
   Begin VB.Image WindowBgImage 
      Height          =   7860
      Left            =   1920
      Picture         =   "Form1.frx":2DEA0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   7635
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

Dim AniFlag As Integer
Dim AniCnt As Integer
'过渡动画

 
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
    
    AniFlag = 2
End Sub
'打开指定文件快捷方式

Public Sub PlaceIcon(appid As String, pic As PictureBox)
    Dim read_OK As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("Pinned", appid, "Noting", read2, 256, App.Path & "\AppList.ini")
    Call ShowIcon(App.Path & "\Links\" & read2, pic)
    pic.ToolTipText = Replace(read2, ".lnk", "")
End Sub
'添加Icon

Private Sub AnimationTimer_Timer()
    '过渡动画执行器
    If AniFlag = 1 Then
        If Form1.Top <= Screen.Height - Form1.Height Then
            AniFlag = 0
        End If
        Form1.Top = Form1.Top - 442
    End If
    If AniFlag = 2 Then
        If Form1.Top >= Screen.Height - 300 Then
            AniFlag = 0
            End
        End If
        Form1.Top = Form1.Top + 700
    End If
End Sub

Private Sub Close_Mask_Btn_Left_Click()
    AniFlag = 2
End Sub

Private Sub Close_Mask_Btn_Right_Click()
    AniFlag = 2
End Sub

Private Sub Close_Mask_Btn_Top_Click()
    AniFlag = 2
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
    '透明窗体
    
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1
    '置顶窗体
    
    Me.Left = (Screen.Width - Form1.Width) / 2
    'Me.Top = Screen.Height - 550 - Form1.Height
    Me.Width = Screen.Width - Me.Left - 50
    Close_Mask_Btn_Right.Width = Me.Width - Close_Mask_Btn_Right.Left - 50
    '校正窗口和控件
    
    
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
    '调用函数，按照AppList.ini中的内容，将快捷方式放置在窗体控件上
    
    Me.Top = Screen.Height
    AniFlag = 1
    '启用动画

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
