VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmLogin_Lib 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   $"frmLogin_Lib.frx":0000
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   1515
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin_Lib.frx":00D4
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmLogin_Lib.frx":099E
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frameupdatetotalsublogin 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Change Max Sub Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   0
      TabIndex        =   40
      Top             =   6960
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtSubLoginMax_ChangeMaxSubLogin 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   45
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtPasswordChangeMaxSubLogin 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   44
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtAdminNameChangeMaxSubLogin 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   43
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdMaxSubLogin 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         Picture         =   "frmLogin_Lib.frx":23B5C0
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "OK"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdunload 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "OK"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change Max Sub Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   1560
         TabIndex        =   49
         Top             =   0
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BorderColor     =   &H00FF8080&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   5775
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sub Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   360
         TabIndex        =   48
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   360
         TabIndex        =   47
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   360
         TabIndex        =   46
         Top             =   840
         Width           =   825
      End
   End
   Begin VB.Frame FrameAdminLogin 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   3240
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtAdminNameL 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         TabIndex        =   52
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtPasswordL 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   51
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdOkL 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Picture         =   "frmLogin_Lib.frx":23BA02
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "OK"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   480
         TabIndex        =   55
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   480
         TabIndex        =   54
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   240
         Left            =   480
         TabIndex        =   53
         Top             =   360
         Width           =   3690
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FF8080&
         Height          =   2295
         Left            =   120
         Top             =   120
         Width           =   4215
      End
      Begin VB.Image Image7 
         Height          =   9645
         Left            =   0
         Picture         =   "frmLogin_Lib.frx":23BE44
         Top             =   0
         Width           =   5100
      End
   End
   Begin VB.Frame frameUSerPasswordGeneration 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "User Password Genetation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   5880
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtREPasswordG 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtPasswordG 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   33
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtUserNameG 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2520
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdOkG 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Picture         =   "frmLogin_Lib.frx":2DC07A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "OK"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2880
         Picture         =   "frmLogin_Lib.frx":2DC4BC
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Exit"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cmbPriv 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF8080&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User Password Genetation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "ReType Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Privilages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      Picture         =   "frmLogin_Lib.frx":2DC8FE
      ScaleHeight     =   1215
      ScaleWidth      =   15480
      TabIndex        =   26
      Top             =   9720
      Width           =   15480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -120
      Picture         =   "frmLogin_Lib.frx":340A44
      ScaleHeight     =   975
      ScaleWidth      =   15480
      TabIndex        =   25
      Top             =   0
      Width           =   15480
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   -360
         X2              =   15960
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   7560
      Width           =   5655
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         Picture         =   "frmLogin_Lib.frx":3A4B8A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exit"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdateTotalSubLogin 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   3000
         Picture         =   "frmLogin_Lib.frx":3A4FCC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Change Max Sub Login"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1560
         Picture         =   "frmLogin_Lib.frx":3A540E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Change Password"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTrapeDoor 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   8520
      Picture         =   "frmLogin_Lib.frx":3A5850
      ScaleHeight     =   2175
      ScaleWidth      =   5055
      TabIndex        =   13
      Top             =   1920
      Width           =   5055
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   3960
         MaskColor       =   &H00808080&
         Picture         =   "frmLogin_Lib.frx":445A86
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "OK"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtReTypePassword 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   2775
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FF8080&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   2160
         Picture         =   "frmLogin_Lib.frx":4466F8
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Labe1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "User &Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "&ReType Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1545
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   10440
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin_Lib.frx":446876
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin_Lib.frx":446CC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6600
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame framenewAdmin 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "New Administrator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C3676&
      Height          =   2175
      Left            =   7920
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtTotalSubLogin 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtNEwReTypePassword 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtNewPassword 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtNewName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdChangeNew 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3360
         Picture         =   "frmLogin_Lib.frx":44711A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Change Password"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdOkNew 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Picture         =   "frmLogin_Lib.frx":44755C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "OK"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New Administrator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label lblTotalSubLogin 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sub Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7C9D2&
         BackStyle       =   0  'Transparent
         Caption         =   "ReType Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006C3676&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   720
      TabIndex        =   9
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adminstrator"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1320
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   103
      ImageHeight     =   97
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":44799E
            Key             =   "admin Login"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":44F028
            Key             =   "afteradmin"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":455C02
            Key             =   "afterallotuser"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":45C544
            Key             =   "afterchangeadminpass"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":462F56
            Key             =   "aftersave"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":463CF0
            Key             =   "aftersubuser"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":46A702
            Key             =   "afteruserlog"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":471654
            Key             =   "Allotuser"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":479086
            Key             =   "changeadminpass"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":480AB8
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":48170A
            Key             =   "subuser"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin_Lib.frx":489190
            Key             =   "userlog"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   14280
      TabIndex        =   58
      Top             =   9360
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   14280
      Picture         =   "frmLogin_Lib.frx":490BC2
      Top             =   8760
      Width           =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   15240
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   960
      Y2              =   9720
   End
   Begin VB.Image Image5 
      Height          =   1500
      Left            =   1920
      Picture         =   "frmLogin_Lib.frx":491804
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Image Image4 
      Height          =   1500
      Left            =   120
      Picture         =   "frmLogin_Lib.frx":499226
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Image Image3 
      Height          =   1485
      Left            =   3840
      Picture         =   "frmLogin_Lib.frx":4A0C48
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1500
      Left            =   1920
      Picture         =   "frmLogin_Lib.frx":4A86BE
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Image imgadmin 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   120
      Picture         =   "frmLogin_Lib.frx":4B00E0
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuUSerTotal 
         Caption         =   "&Change Total Sub User"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuAllotPassword 
         Caption         =   "Allot &Password to User"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "frmLogin_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c1 As New ClassMaster
Dim classED As New clsEncodeDecode
Private ChangeFlag, loginfloag As Integer
Dim flagAdmin, flagUser As Boolean
Dim rsAdminLogin, rsUserLogin As Recordset
Dim flagValidMenu As Boolean, counter As Integer
'Dim con As Connection

Public Sub AddLogin()
Dim rsrecordlogin As Recordset
Set rsrecordlogin = New Recordset
End Sub

Private Sub cmdChangePass_Click()
 If flagAdmin = True Then
  
 End If
End Sub
Sub proLogerNameUser()

 MDIForm1.mnuLoginName.Caption = "User Login " + "[ " + UCase(txtName.Text) + " ]"

End Sub

Sub proLogerNameAdmin()

 MDIForm1.mnuLoginName.Caption = "Admin Login " + "[ " + UCase(txtName.Text) + " ]"

End Sub
Sub proAdminLogin()
If ChangeFlag = 0 Then
   rsAdminLogin.Close
  rsAdminLogin.Open "select * from Admin_Login_master where company_Id=" & Global_CompanyId & " ", con, adOpenKeyset, adLockOptimistic
  If rsAdminLogin.RecordCount = 0 Then
    Exit Sub
  End If
   If classED.DecodeString(rsAdminLogin.Fields(0).Value) = Trim(txtName.Text) And classED.DecodeString(rsAdminLogin.Fields(1).Value) = Trim(txtPassword.Text) Then
   Call proLogerNameAdmin 'For Displaying Login Name
    'MsgBox ("Password Succeed")
    Unload Me
    Call c1.proAdminLogin
    
    loginflag = 1
     
    Exit Sub
   Else
   If counter >= 3 Then
     End  'After three attempt software will be terminated
   End If
    counter = counter + 1 ' for maximum trying of unknown login
    MsgBox "ACCESS DENIED DUE TO SECURITY", vbOKOnly, "WORNG PASSWORD"
    txtName.Text = ""
    txtPassword.Text = ""
    txtName.SetFocus
    Exit Sub
   End If
ElseIf ChangeFlag = 1 Then
    If classED.DecodeString(rsAdminLogin.Fields(0).Value) = Trim(txtName.Text) And classED.DecodeString(rsAdminLogin.Fields(1).Value) = txtPassword.Text Then
     Label1.Caption = "Enter New Admin Name"
     Label2.Caption = "Enter New Password"
     Label3.Visible = True
     txtReTypePassword.Visible = True
     txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     ChangeFlag = 2
     Else
       MsgBox "Access Denied Due to Security", vbInformation + vbOKOnly, "Friendly Advise"
       
     End If
ElseIf ChangeFlag = 2 Then
    If txtPassword.Text = "" Then
      MsgBox "You Must Not Let Password To Be Blank", vbInformation + vbOKOnly, "Friendly Advice"
      txtPassword.SetFocus
      Exit Sub
      ElseIf txtName.Text = "" Then
       MsgBox "You Must Not Let User Name To Be Blank", vbInformation + vbOKOnly, "friendly Advice"
       txtName.SetFocus
       Exit Sub
    Else
    If txtPassword.Text = txtReTypePassword.Text Then
      
      Dim strsql As String
      strsql = "update Admin_login_master set admin_name='" & classED.EncodeString(Trim(txtName)) & "',password1='" & classED.EncodeString(Trim(txtPassword)) & "' where company_Id=" & Val(Global_CompanyId) & " "
      'MsgBox strsql
      con.Execute (strsql)
      txtReTypePassword.Text = ""
     txtReTypePassword.Visible = False
     Label3.Visible = False
     Label1.Caption = "Enter Admin Name"
     Label2.Caption = "Enter Password"
     txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     ChangeFlag = 0
     Else
      MsgBox "Enter Same password in retype password"
      txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     End If
     Exit Sub
End If
End If

End Sub
Function ProGoToRecord() As Boolean
If rsUserLogin.RecordCount = 0 Then
MsgBox "User Does Not Exist", vbInformation + vbOKOnly, "EasyLib"
 ProGoToRecord = False
Exit Function
End If
rsUserLogin.MoveFirst
While Not rsUserLogin.EOF
     If classED.DecodeString(Trim(rsUserLogin.Fields(0).Value)) = Trim(txtName.Text) And rsUserLogin.Fields(2).Value = "& Global_CompanyId  &" Then
       MsgBox "Succed"
       ProGoToRecord = True
       Exit Function
     End If
     rsUserLogin.MoveNext
     
Wend
ProGoToRecord = False
End Function

Sub proUserLogin()
'If ProGoToRecord = True Then
 If ChangeFlag = 0 Then
 '-----------------------Selection Of User---------------------
  Dim Rs As Recordset
 Set Rs = New Recordset
 Rs.Open "select * from user_Login where company_Id=" & Global_CompanyId & " and user_Name='" & classED.EncodeString(Trim(txtName.Text)) & "' and Password1='" & classED.EncodeString(Trim(txtPassword.Text)) & "'", con, adOpenKeyset, adLockOptimistic
 If Rs.RecordCount <= 0 Then
  If counter >= 3 Then
     End 'After three attempt software will be terminated
   End If
    counter = counter + 1 ' for maximum trying of unknown login
    MsgBox "User Does Not Exist", vbInformation + vbOKOnly
    txtPassword.Text = ""
    txtName.Text = ""
    Exit Sub
 End If
'=================================================================
   
   '--------------For Searching The Records------------------
   If classED.DecodeString(Rs.Fields(0).Value) = Trim(txtName.Text) And classED.DecodeString(Rs.Fields(1).Value) = Trim(txtPassword.Text) Then
    Call proLogerNameUser  'For Displaying Login Name
    'MsgBox ("Password Succeed")
    Unload Me
    Call c1.proUserLogin 'Authority of User

    loginflag = 1
    Exit Sub
   Else
    MsgBox "ACCESS DENIED DUE TO SECURITY", vbOKOnly, "WORNG PASSWORD"
    txtName.Text = ""
    txtPassword.Text = ""
    txtName.SetFocus
    Exit Sub
   End If
ElseIf ChangeFlag = 1 Then
'-----------------------Selection Of User---------------------
  Dim rs1 As Recordset
 Set rs1 = New Recordset
 rs1.Open "select * from user_Login where company_Id=" & Global_CompanyId & " and user_Name='" & classED.EncodeString(Trim(txtName.Text)) & "' and Password1='" & classED.EncodeString(Trim(txtPassword.Text)) & "'", con, adOpenKeyset, adLockOptimistic
 If rs1.RecordCount <= 0 Then
    MsgBox "User Does Not Exist", vbInformation + vbOKOnly
    txtPassword.Text = ""
    txtName.Text = ""
    Exit Sub
 End If
'=================================================================
    If classED.DecodeString(rs1.Fields(0).Value) = Trim(txtName.Text) And classED.DecodeString(rs1.Fields(1).Value) = Trim(txtPassword.Text) Then
     Label1.Caption = "Enter New User Name"
     Label2.Caption = "Enter New Password"
     Label3.Visible = True
     txtReTypePassword.Visible = True
     'txtRePassword.Visible = True
     txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     ChangeFlag = 2
     Else
       MsgBox "Access Denied Due to Security!!!!", vbInformation + vbOKOnly, "Friendly Advise"
       
     End If
ElseIf ChangeFlag = 2 Then
'-----------------------Selection Of User---------------------
'  Dim Rs2 As Recordset
' Set Rs2 = New Recordset
' Rs2.Open "select * from user_Login where company_Id=" & Global_CompanyId & " and user_Name='" & classED.EncodeString(Trim(txtName.Text)) & "' and Password1='" & classED.EncodeString(Trim(txtPassword.Text)) & "'", con, adOpenKeyset, adLockOptimistic
' If Rs2.RecordCount <= 0 Then
'    MsgBox "User Does Not Exist", vbInformation + vbOKOnly
'    txtPassword.Text = ""
'    txtName.Text = ""
'    Exit Sub
' End If
'=================================================================
    If Trim(txtPassword.Text) = "" Then
      MsgBox "You Must Not Let Password To Be Blank", vbInformation + vbOKOnly, "Friendly Advice"
      txtPassword.SetFocus
      Exit Sub
      ElseIf Trim(txtName.Text) = "" Then
       MsgBox "You Must Not Let User Name To Be Blank", vbInformation + vbOKOnly, "friendly Advice"
       txtName.SetFocus
       Exit Sub
    Else
    If Trim(txtPassword.Text) = Trim(txtReTypePassword.Text) Then
      
      Dim strsql As String
      strsql = "update User_login set user_name='" & classED.EncodeString(Trim(txtName)) & "',password1='" & classED.EncodeString(Trim(txtPassword)) & "' where company_Id=" & Val(Global_CompanyId) & " and User_name='" & classED.EncodeString(Trim(txtName.Text)) & "' "
     ' MsgBox strsql
      con.Execute (strsql)
      txtReTypePassword.Text = ""
     txtReTypePassword.Visible = False
     Label3.Visible = False
     Label1.Caption = "Enter User Name"
     Label2.Caption = "Enter Password"
     txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     ChangeFlag = 0
     Else
      MsgBox "Enter Same password in retype password"
      txtName.Text = ""
     txtPassword.Text = ""
     txtName.SetFocus
     End If
     Exit Sub
End If
'End If
End If
End Sub

Private Sub CmdExit_Click()
frameUSerPasswordGeneration.Visible = False
frameUSerPasswordGeneration.Top = 4440
End Sub

Private Sub cmdMaxSubLogin_Click()
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select * from Admin_login_master where Admin_name='" & classED.EncodeString(Trim(txtAdminNameChangeMaxSubLogin)) & "' and password1='" & classED.EncodeString(Trim(txtPasswordChangeMaxSubLogin)) & "' and company_Id=" & Global_CompanyId & "", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
  MsgBox "Administrator Does Not Exist", vbInformation + vbOKOnly
  
Else
   If checkValidMax = True Then 'for not going to Under Limit
  strsql = "update admin_login_master set total_Sub_login=" & Val(txtSubLoginMax_ChangeMaxSubLogin) & " where Admin_name='" & classED.EncodeString(Trim(txtAdminNameChangeMaxSubLogin)) & "' and password1='" & classED.EncodeString(Trim(txtPasswordChangeMaxSubLogin)) & "' and company_Id=" & Global_CompanyId & ""
 ' MsgBox strsql
  con.Execute (strsql)
  Else
    MsgBox "Change is not Allowed due to Already Alloted Limited SubUser ID", vbInformation + vbOKOnly, "Friendly Advice"
  End If
End If
 frameupdatetotalsublogin.Enabled = False
 frameupdatetotalsublogin.Visible = False
 cmdUpdateTotalSubLogin.Enabled = True
 txtAdminNameChangeMaxSubLogin.Text = ""
 txtPasswordChangeMaxSubLogin.Text = ""
 txtSubLoginMax_ChangeMaxSubLogin.Text = ""
End Sub
Private Function checkValidMax() As Boolean
 
 If rsUserLogin.RecordCount < Val(txtSubLoginMax_ChangeMaxSubLogin.Text) Then
    checkValidMax = True
 Else
   checkValidMax = False
 End If
End Function
Private Sub cmdOk_Click()
'Unload Me
'Load MDIForm1
'MDIForm1.Show
If flagAdmin = True Then
 Call proAdminLogin
 End If
 If flagUser = True Then
  Call proUserLogin
 End If
End Sub
Private Function validDataAdmin() As Boolean
  If txtNewName = "" Then
     MsgBox "Enter New Administrator Name", vbInformation + vbOKOnly
     txtNewName.SetFocus
     validDataAdmin = False
  ElseIf txtNewPassword = "" Then
    MsgBox "Enter Password of Administartor", vbInformation + vbOKOnly
    txtNewPassword.SetFocus
    validDataAdmin = False
  ElseIf txtNEwReTypePassword = "" Then
       MsgBox "Enter RePassword of Administartor", vbInformation + vbOKOnly
       txtNEwReTypePassword.SetFocus
       validDataAdmin = False
  ElseIf txtTotalSubLogin = "" Then
        MsgBox "Enter Total Sub Logins of Administartor", vbInformation + vbOKOnly
        txtTotalSubLogin.SetFocus
        validDataAdmin = False
  End If
  If txtNEwReTypePassword.Text = txtNewPassword.Text Then
       validDataAdmin = True
   Else
     MsgBox "Password Should be correct in both Password Boxes", vbInformation + vbOKOnly
     validDataAdmin = False
      txtNewPassword.Text = ""
      txtNEwReTypePassword.Text = ""
      txtNewPassword.SetFocus
  End If
End Function
Sub proNewAdmin()
lblTotalSubLogin.Visible = True
txtTotalSubLogin.Visible = True
lblTotalSubLogin.Enabled = True
txtTotalSubLogin.Enabled = True
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select * from Admin_login_master where company_Id=" & Global_CompanyId & " ", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
    If validDataAdmin = True Then
         strsql = "insert into Admin_login_master values('" & classED.EncodeString(Trim(txtNewName.Text)) & "','" & classED.EncodeString(Trim(txtNewPassword)) & "'," & Global_CompanyId & "," & Val(txtTotalSubLogin) & ")"
         con.Execute (strsql)
         
    Else
         Exit Sub
    End If
Else
  MsgBox "Administrator Login Already Exist", vbInformation + vbOKOnly
End If
         framenewAdmin.Visible = False
         framenewAdmin.Enabled = False
         framenewAdmin.Top = 3480
         framenewAdmin.Left = 480
End Sub

Private Sub cmdOk_GotFocus()
cmdOK.BackColor = &H80000003
End Sub

Private Sub cmdOk_LostFocus()
cmdOK.BackColor = &HC0C0C0
End Sub

Private Sub Cmdok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOK.SetFocus
End Sub

Private Sub cmdOkG_Click()
If txtUserNameG.Text = "" Then
   MsgBox "Input New User Name", vbInformation + vbOKOnly, "Freindly Advice"
   txtUserNameG.SetFocus
ElseIf txtPasswordG = "" Then
   MsgBox "Input Password", vbInformation + vbOKOnly, "Freindly Advice"
   txtPasswordG.SetFocus
   Call ProSetAdminUser
ElseIf txtREPasswordG.Text = "" Then
    MsgBox "ReType Password", vbInformation + vbOKOnly, "Freindly Advice"
    txtREPasswordG.SetFocus
 ElseIf txtPasswordG.Text = txtREPasswordG.Text Then
       Call AllotPassword
 Else
      MsgBox "Input Password Correctly", vbInformation + vbOKOnly, "Freindly Advice"
      txtREPasswordG.Text = ""
      txtPasswordG = ""
      txtPasswordG.SetFocus
 End If
 flagAdmin = True
End Sub

Private Sub cmdOkL_Click()
If txtAdminNameL = "" Then
 MsgBox "Input Admin Name", vbInformation + vbOKOnly, "friendly Advice"
 txtAdminNameL.SetFocus
ElseIf txtPasswordL.Text = "" Then
  MsgBox "Password Can not Be Blank", vbInformation + vbOKOnly, "Friendly Advice"
  txtPasswordL.SetFocus
Else
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select Admin_name,Password1 from Admin_Login_Master where company_Id=" & Global_CompanyId & "", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
'If Rs.RecordCount <= 0 Then
  If counter >= 3 Then
     End 'After three attempt software will be terminated
   End If
    counter = counter + 1 ' for maximum trying of unknown login
    MsgBox "Administrator Login Does Not Exist", vbInformation + vbOKOnly, "Friendly Advice"
    txtAdminNameL.Text = ""
    txtPasswordL.Text = ""
Else
    If Rs.Fields(0).Value = classED.EncodeString(Trim(txtAdminNameL.Text)) And Rs.Fields(1).Value = classED.EncodeString(Trim(txtPasswordL.Text)) Then
      'MsgBox "password Succeed", vbInformation + vbOKOnly
     'sdfsdfsfsdfsdfsfs
     If flagValidMenu = True Then
        frameupdatetotalsublogin.Visible = True
        frameupdatetotalsublogin.Top = 2200
        frameupdatetotalsublogin.Left = 8500
        cmdUpdateTotalSubLogin.Enabled = False
        frameupdatetotalsublogin.Enabled = True
     Else
        frameUSerPasswordGeneration.Top = 2200
        frameUSerPasswordGeneration.Left = 8500
        frameUSerPasswordGeneration.Enabled = True
        frameUSerPasswordGeneration.Visible = True
     End If
      
       
    'dsfsdfsdfsdfsdfsdf
      txtAdminNameL.Text = ""
      txtPasswordL.Text = ""
     FrameAdminLogin.Visible = False
     Else
     MsgBox "Invalid Login Name And Password", vbInformation + vbOKOnly, "Friendly Advice"
    End If
End If
End If
End Sub

Private Sub cmdOkNew_Click()
If flagAdmin = True Then
  Call proNewAdmin
End If
If flagUser = True Then
  'Call proNewUser
End If
End Sub

Private Sub cmdTrapeDoor_Click()
If flagAdmin = True Then
If txtName = "" Then
  MsgBox "Input Admin Name", vbInformation + vbOKOnly
  txtName.SetFocus
ElseIf txtPassword.Text = "" Then
 MsgBox "input Password", vbInformation + vbOKOnly
 txtPassword.SetFocus
Else
strsql = "update Admin_login_master set admin_name='" & classED.EncodeString(Trim(txtName)) & "',password1='" & classED.EncodeString(Trim(txtPassword)) & "' where company_Id=" & Global_CompanyId & ""
'MsgBox strsql
con.Execute (strsql)
strsql = "delete from user_login where company_Id=" & Global_CompanyId & ""
con.Execute (strsql)
End If
End If
End Sub



Private Sub cmdUpdate_Click()
a = MsgBox("Do you Want of Change Password", vbOKCancel + vbQuestion, "Change Password")
Select Case a
 Case vbOK
 Picture2.Visible = True
  Label2.Caption = "Old Password"
  If flagAdmin = True Then
   Label1.Caption = "Old Admin Name"
  ElseIf flagUser = True Then
     Label1.Caption = "Old User Name"
  End If
  cmdUpdate.Enabled = False
  ChangeFlag = 1
 Case vbCancel
  ChangeFlag = 0
 cmdUpdate.Enabled = True
End Select

End Sub

Private Sub cmdUpdate_GotFocus()
cmdUpdate.BackColor = &H80000003
End Sub

Private Sub cmdUpdate_LostFocus()
cmdUpdate.BackColor = &HC0C0C0
End Sub

Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdUpdate.SetFocus
End Sub

Private Sub cmdUpdateTotalSubLogin_Click()
frameupdatetotalsublogin.Visible = True
frameupdatetotalsublogin.Top = 600
frameupdatetotalsublogin.Left = 1560
cmdUpdateTotalSubLogin.Enabled = False
frameupdatetotalsublogin.Enabled = True
End Sub

Private Sub cmdUpdateTotalSubLogin_GotFocus()
cmdUpdateTotalSubLogin.BackColor = &H80000003
End Sub

Private Sub cmdUpdateTotalSubLogin_LostFocus()
cmdUpdateTotalSubLogin.BackColor = &HC0C0C0
End Sub

Private Sub cmdUpdateTotalSubLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdUpdateTotalSubLogin.SetFocus
End Sub

Private Sub Command1_Click()
      strsql = "delete  admin_login_master where company_id=1"
      con.Execute (strsql)
End Sub




Private Sub Command2_Click()
If MsgBox("Do You Want To Exit", vbQuestion + vbYesNo, "Visitor Managers") = vbYes Then
    End
End If
End Sub

Private Sub Command2_GotFocus()
Command2.BackColor = &H80000003
End Sub

Private Sub Command2_LostFocus()
Command2.BackColor = &HC0C0C0
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Private Sub Form_Load()
Call addpriv
Frame3.Visible = False
Picture2.Visible = False
Call AddLogin
Set rsAdminLogin = New Recordset
Call ProSetAdminUser
flagAdmin = True
End Sub
Sub ProSetAdminUser()
rsAdminLogin.Open "select * from Admin_Login_master where company_Id=" & Global_CompanyId & " ", con, adOpenKeyset, adLockOptimistic
If rsAdminLogin.RecordCount = 0 Then
  framenewAdmin.Left = 600
  framenewAdmin.Top = 840
   framenewAdmin.Visible = True
   framenewAdmin.Enabled = True
  ' txtNewName.SetFocus
End If
Set rsUserLogin = New Recordset
rsUserLogin.Open "select * from User_Login where company_Id=" & Global_CompanyId & " ", con, adOpenKeyset, adLockOptimistic
Labe1.Caption = "Admin &Name"

End Sub
Sub addpriv()
Dim rsd As Recordset
Set rsd = New Recordset
rsd.Open "select priv from priv_master", con, adOpenKeyset, adLockOptimistic
rsd.MoveFirst
Do While Not rsd.EOF = True
cmbPriv.AddItem (rsd!priv)
rsd.MoveNext
Loop
rsd.Close
End Sub
Sub AllotPassword()
Dim total, remen As Integer
total = Val(rsAdminLogin.Fields(3).Value)
'MsgBox "Total Users is:" & total

If total > rsUserLogin.RecordCount Then
   strsql = "insert into User_Login values('" & classED.EncodeString(Trim(txtUserNameG.Text)) & "','" & classED.EncodeString(Trim(txtPasswordG.Text)) & "'," & Global_CompanyId & ",'" & Trim(cmbPriv.Text) & "')"
   'MsgBox strsql
   con.Execute (strsql)
   remen = total - Val(rsUserLogin.RecordCount) - 1
   MsgBox "User created Successfully,remaing Users:" & remen & "", vbInformation
   rsAdminLogin.Close
   txtUserNameG.Text = ""
   txtPasswordG.Text = ""
   txtREPasswordG.Text = ""
'   cmbPriv.Text = ""
   txtUserNameG.SetFocus
   Call ProSetAdminUser
Else
MsgBox "Users can not be more than " & total & " Limit", vbInformation + vbOKOnly, "Friendly Advice"
frameUSerPasswordGeneration.Left = 6240
frameUSerPasswordGeneration.Top = 4920
   rsAdminLogin.Close
   Call ProSetAdminUser
End If
  

  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image6.Picture = LoadPicture(App.Path & "\Support\stop.bmp")
Label18.ForeColor = &HFF8080
End Sub

Private Sub Form_Resize()
'frmLogin_Lib.ScaleWidth = 7125
'frmLogin_Lib.Width = 7245
frmLogin_Lib.ScaleMode = False
If frmLogin_Lib.Height > 4215 Then
'frmLogin_Lib.Height = 4215
End If
End Sub



Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SetFocus
End Sub

Private Sub Image2_Click()
TabStrip1.Tabs(2).Selected = True
Picture2.Visible = True
imgadmin.Picture = ImageList2.ListImages(1).Picture
'imgadmin.Picture = LoadPicture(App.Path + "\Support\afteradmin.bmp")
'............other images.............
 Image2.Picture = ImageList2.ListImages(7).Picture
 Image3.Picture = ImageList2.ListImages(11).Picture
 Image4.Picture = ImageList2.ListImages(8).Picture
 Image5.Picture = ImageList2.ListImages(9).Picture
'Image2.Picture = LoadPicture(App.Path + "\Support\afteruserlog.bmp")
'............other images.............
'Image3.Picture = LoadPicture(App.Path + "\Support\subuser.bmp")
'Image4.Picture = LoadPicture(App.Path + "\Support\allotuser.bmp")
'Image5.Picture = LoadPicture(App.Path + "\Support\Changeadminpass.bmp")
'imgadmin.Picture = LoadPicture(App.Path + "\Support\admin Login.bmp")
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image2.BorderStyle = 1
End Sub

Private Sub Image3_Click()
Picture2.Visible = False
Call mnuUSerTotal_Click
imgadmin.Picture = ImageList2.ListImages(1).Picture
'imgadmin.Picture = LoadPicture(App.Path + "\Support\afteradmin.bmp")
'............other images.............
 Image2.Picture = ImageList2.ListImages(12).Picture
 Image3.Picture = ImageList2.ListImages(6).Picture
 Image4.Picture = ImageList2.ListImages(8).Picture
 Image5.Picture = ImageList2.ListImages(9).Picture
'Image3.Picture = LoadPicture(App.Path + "\Support\aftersubuser.bmp")
'............other images.............
'Image2.Picture = LoadPicture(App.Path + "\Support\userlog.bmp")
'imgadmin.Picture = LoadPicture(App.Path + "\Support\admin login.bmp")
'Image4.Picture = LoadPicture(App.Path + "\Support\allotuser.bmp")
'Image5.Picture = LoadPicture(App.Path + "\Support\Changeadminpass.bmp")
End Sub

Private Sub Image4_Click()
Picture2.Visible = False
Call mnuAllotPassword_Click
imgadmin.Picture = ImageList2.ListImages(1).Picture
'imgadmin.Picture = LoadPicture(App.Path + "\Support\afteradmin.bmp")
'............other images.............
 Image2.Picture = ImageList2.ListImages(12).Picture
 Image3.Picture = ImageList2.ListImages(11).Picture
 Image4.Picture = ImageList2.ListImages(3).Picture
 Image5.Picture = ImageList2.ListImages(9).Picture
'Image4.Picture = LoadPicture(App.Path + "\Support\afterallotuser.bmp")
'............other button images.............
'Image2.Picture = LoadPicture(App.Path + "\Support\userlog.bmp")
'Image3.Picture = LoadPicture(App.Path + "\Support\subuser.bmp")
'imgadmin.Picture = LoadPicture(App.Path + "\Support\admin login.bmp")
'Image5.Picture = LoadPicture(App.Path + "\Support\Changeadminpass.bmp")
End Sub

Private Sub Image5_Click()
Call cmdUpdate_Click
FrameAdminLogin.Visible = False
imgadmin.Picture = ImageList2.ListImages(1).Picture
'imgadmin.Picture = LoadPicture(App.Path + "\Support\afteradmin.bmp")
'............other images.............
 Image2.Picture = ImageList2.ListImages(12).Picture
 Image3.Picture = ImageList2.ListImages(11).Picture
 Image4.Picture = ImageList2.ListImages(8).Picture
 Image5.Picture = ImageList2.ListImages(4).Picture
'Image5.Picture = LoadPicture(App.Path + "\Support\afterChangeadminpass.bmp")
'............other images.............
'Image2.Picture = LoadPicture(App.Path + "\Support\userlog.bmp")
'Image3.Picture = LoadPicture(App.Path + "\Support\subuser.bmp")
'Image4.Picture = LoadPicture(App.Path + "\Support\allotuser.bmp")
'imgadmin.Picture = LoadPicture(App.Path + "\Support\admin Login.bmp")
End Sub

Private Sub Image6_Click()
Call Command2_Click
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image6.Picture = LoadPicture(App.Path & "\Support\aftersave.bmp")
Label18.ForeColor = &H8080FF
End Sub

Private Sub imgadmin_Click()

TabStrip1.Tabs(1).Selected = True
FrameAdminLogin.Visible = False
Picture2.Visible = True
txtName.SetFocus


imgadmin.Picture = ImageList2.ListImages(2).Picture
'imgadmin.Picture = LoadPicture(App.Path + "\Support\afteradmin.bmp")
'............other images.............
 Image2.Picture = ImageList2.ListImages(12).Picture
 Image3.Picture = ImageList2.ListImages(11).Picture
 Image4.Picture = ImageList2.ListImages(8).Picture
 Image5.Picture = ImageList2.ListImages(9).Picture
'Image2.Picture = LoadPicture(App.Path + "\Support\userlog.bmp")
'Image3.Picture = LoadPicture(App.Path + "\Support\subuser.bmp")
'Image4.Picture = LoadPicture(App.Path + "\Support\allotuser.bmp")
'Image5.Picture = LoadPicture(App.Path + "\Support\Changeadminpass.bmp")
End Sub

Private Sub imgadmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'imgadmin.BorderStyle = 1
End Sub

Private Sub mnuAllotPassword_Click()
flagValidMenu = False
FrameAdminLogin.Visible = True
FrameAdminLogin.Top = 2200
FrameAdminLogin.Left = 8500
txtAdminNameL.SetFocus
End Sub

Private Sub mnuUSerTotal_Click()
flagValidMenu = True
FrameAdminLogin.Visible = True
FrameAdminLogin.Top = 2200
FrameAdminLogin.Left = 8500
txtAdminNameL.SetFocus
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
   flagAdmin = True
   flagUser = False
   cmdUpdateTotalSubLogin.Enabled = True
   Labe1.Caption = "Admin &Name"
End If
If TabStrip1.SelectedItem.Index = 2 Then
   flagUser = True
   flagAdmin = False
   cmdUpdateTotalSubLogin.Enabled = False
   Labe1.Caption = "User &Name"
End If

End Sub

Private Sub TabStrip1_KeyPress(KeyAscii As Integer)
If TabStrip1.SelectedItem.Index = 1 Then
If KeyAscii = 1 Then
  cmdTrapeDoor.Enabled = True
  cmdTrapeDoor.Visible = True
  Command1.Visible = True
  Command1.Enabled = True
End If
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdOk_Click
End If
End Sub

