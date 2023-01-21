VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select File"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
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
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.FileListBox FileSelect 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   2880
      Pattern         =   "*.exe"
      TabIndex        =   2
      Top             =   600
      Width           =   2640
   End
   Begin VB.DirListBox DirSelect 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox DriveSelect 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdOk_Click()

OkClick = True
Me.Hide

End Sub

Private Sub DirSelect_Change()

FileSelect.Path = DirSelect.Path

End Sub

Private Sub DriveSelect_Change()

On Error GoTo ErrorHandler

DirSelect.Path = DriveSelect.Drive

ErrorHandler:
If Err.Number = 68 Then MsgBox "The device is not ready", vbOK, "Error!"

End Sub

Private Sub Form_Load()

DirSelect.Path = "c:\"

End Sub

