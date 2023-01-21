VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2580
      TabIndex        =   51
      Top             =   5820
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "IWADS"
      TabPicture(0)   =   "frmOptionsTabs.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtStrife"
      Tab(0).Control(1)=   "cmdStrife"
      Tab(0).Control(2)=   "cmdHexen"
      Tab(0).Control(3)=   "txtHexen"
      Tab(0).Control(4)=   "txtDoom"
      Tab(0).Control(5)=   "cmdDoom"
      Tab(0).Control(6)=   "cmdSaveIwadData"
      Tab(0).Control(7)=   "cmdDoom2"
      Tab(0).Control(8)=   "txtDoom2"
      Tab(0).Control(9)=   "cmdPlutonia"
      Tab(0).Control(10)=   "txtPlutonia"
      Tab(0).Control(11)=   "cmdTnt"
      Tab(0).Control(12)=   "txtTnt"
      Tab(0).Control(13)=   "cmdHeretic"
      Tab(0).Control(14)=   "txtFreeDoom"
      Tab(0).Control(15)=   "cmdFreeDoom"
      Tab(0).Control(16)=   "txtHeretic"
      Tab(0).Control(17)=   "Label1(6)"
      Tab(0).Control(18)=   "Label1(5)"
      Tab(0).Control(19)=   "Label1(0)"
      Tab(0).Control(20)=   "Label1(1)"
      Tab(0).Control(21)=   "Label1(2)"
      Tab(0).Control(22)=   "Label1(3)"
      Tab(0).Control(23)=   "Label1(4)"
      Tab(0).Control(24)=   "Label2"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "EXE"
      TabPicture(1)   =   "frmOptionsTabs.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtGzdoom"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdGzdoom"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdZdoom"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtZdoom"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdChocolate"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtChocolate"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdEdge"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtEdge"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmdSkulltag"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtSkulltag"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdprBoom"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtprBoom"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdEternity"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdVavoom"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtEternity"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtVavoom"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmdSaveExe"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).ControlCount=   25
      Begin VB.CommandButton cmdSaveExe 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   50
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtVavoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   4380
         Width           =   2775
      End
      Begin VB.TextBox txtEternity 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton cmdVavoom 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   44
         Top             =   4380
         Width           =   1815
      End
      Begin VB.CommandButton cmdEternity 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   43
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtprBoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3300
         Width           =   2775
      End
      Begin VB.CommandButton cmdprBoom 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   41
         Top             =   3300
         Width           =   1815
      End
      Begin VB.TextBox txtStrife 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton cmdStrife 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   26
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmdHexen 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   25
         Top             =   3300
         Width           =   1815
      End
      Begin VB.TextBox txtHexen 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3300
         Width           =   2775
      End
      Begin VB.TextBox txtDoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdDoom 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   22
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveIwadData 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72480
         TabIndex        =   21
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdDoom2 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   20
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtDoom2 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1140
         Width           =   2775
      End
      Begin VB.CommandButton cmdPlutonia 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtPlutonia 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmdTnt 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   16
         Top             =   2220
         Width           =   1815
      End
      Begin VB.TextBox txtTnt 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2220
         Width           =   2775
      End
      Begin VB.CommandButton cmdHeretic 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtFreeDoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4380
         Width           =   2775
      End
      Begin VB.CommandButton cmdFreeDoom 
         Caption         =   "Select IWAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73920
         TabIndex        =   12
         Top             =   4380
         Width           =   1815
      End
      Begin VB.TextBox txtHeretic 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72060
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtSkulltag 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CommandButton cmdSkulltag 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtEdge 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2220
         Width           =   2775
      End
      Begin VB.CommandButton cmdEdge 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   7
         Top             =   2220
         Width           =   1815
      End
      Begin VB.TextBox txtChocolate 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmdChocolate 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtZdoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   2775
      End
      Begin VB.CommandButton cmdZdoom 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   3
         Top             =   1140
         Width           =   1815
      End
      Begin VB.CommandButton cmdGzdoom 
         Caption         =   "Select EXE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtGzdoom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   2940
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Vavoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4380
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Eternity Engine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   300
         TabIndex        =   48
         Top             =   3780
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "prBoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Strife"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -74580
         TabIndex        =   40
         Top             =   3840
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hexen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -74640
         TabIndex        =   39
         Top             =   3300
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Doom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74580
         TabIndex        =   38
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Doom II"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -74760
         TabIndex        =   37
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Plutonia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -74760
         TabIndex        =   36
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TNT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -74460
         TabIndex        =   35
         Top             =   2220
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Heretic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74760
         TabIndex        =   34
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "FreeDoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   4380
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Skulltag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   300
         TabIndex        =   32
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "EDGE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   420
         TabIndex        =   31
         Top             =   2220
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Chocolate Doom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   180
         TabIndex        =   30
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ZDoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "GZDoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChocolate_Click()

Call OpenFile("*.exe", "Select Chocolate Executable", txtChocolate)

End Sub

Private Sub cmdDoom_Click()

Call OpenFile("*.wad", "Select Doom IWAD", txtDoom)

End Sub

Private Sub cmdDoom2_Click()

Call OpenFile("*.wad", "Select Doom II IWAD", txtDoom2)

End Sub

Private Sub cmdEdge_Click()

Call OpenFile("*.exe", "Select EDGE Executable", txtEdge)

End Sub

Private Sub cmdEternity_Click()

Call OpenFile("*.exe", "Select Eternity Engine Executable", txtEternity)

End Sub

Private Sub cmdFreeDoom_Click()

Call OpenFile("*.wad", "Select FreeDoom IWAD", txtFreeDoom)

End Sub

Private Sub cmdGzdoom_Click()

Call OpenFile("*.exe", "Select GZDoom Executable", txtGzdoom)

End Sub

Private Sub cmdHeretic_Click()

Call OpenFile("*.wad", "Select Heretic IWAD", txtHeretic)

End Sub

Private Sub cmdHexen_Click()

Call OpenFile("*.wad", "Select Hexen IWAD", txtHexen)

End Sub

Private Sub cmdOk_Click()

Unload Me

End Sub

Private Sub cmdPlutonia_Click()

Call OpenFile("*.wad", "Select Plutonia IWAD", txtPlutonia)

End Sub

Private Sub cmdprBoom_Click()

Call OpenFile("*.exe", "Select prBoom Executable", txtprBoom)

End Sub

Private Sub cmdSaveExe_Click()

rsExe.MoveFirst
Call SaveData("EDIT", txtGzdoom, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtZdoom, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtChocolate, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtEdge, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtSkulltag, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtprBoom, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtEternity, rsExe, "ExePath")
rsExe.MoveNext
Call SaveData("EDIT", txtVavoom, rsExe, "ExePath")
rsExe.MoveFirst


End Sub

Private Sub cmdSaveIwadData_Click()

rsIwads.MoveFirst
Call SaveData("EDIT", txtDoom, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtDoom2, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtTnt, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtPlutonia, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtHeretic, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtHexen, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtStrife, rsIwads, "IwadPath")
rsIwads.MoveNext
Call SaveData("EDIT", txtFreeDoom, rsIwads, "IwadPath")
rsIwads.MoveFirst

End Sub

Private Sub cmdSkulltag_Click()

Call OpenFile("*.exe", "Select Skulltag Executable", txtSkulltag)

End Sub

Private Sub cmdStrife_Click()

Call OpenFile("*.wad", "Select Strife IWAD", txtStrife)

End Sub

Private Sub cmdTnt_Click()

Call OpenFile("*.wad", "Select TNT IWAD", txtTnt)

End Sub

Private Sub cmdVavoom_Click()

Call OpenFile("*.exe", "Select Vavoom Executable", txtVavoom)

End Sub

Private Sub cmdZdoom_Click()

Call OpenFile("*.exe", "Select ZDoom Executable", txtZdoom)

End Sub

Private Sub Form_Load()

rsExe.MoveFirst
Call LoadData(txtGzdoom, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtZdoom, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtChocolate, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtEdge, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtSkulltag, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtprBoom, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtEternity, rsExe, "ExePath")
rsExe.MoveNext
Call LoadData(txtVavoom, rsExe, "ExePath")
rsExe.MoveFirst

rsIwads.MoveFirst
Call LoadData(txtDoom, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtDoom2, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtTnt, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtPlutonia, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtHeretic, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtHexen, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtStrife, rsIwads, "IwadPath")
rsIwads.MoveNext
Call LoadData(txtFreeDoom, rsIwads, "IwadPath")
rsIwads.MoveFirst


End Sub

