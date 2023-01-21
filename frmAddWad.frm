VERSION 5.00
Begin VB.Form frmWadForm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Default Values"
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
      Left            =   1860
      TabIndex        =   22
      Top             =   3900
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   180
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5820
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtNegro 
      Height          =   315
      Left            =   300
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6780
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtDeh2 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1860
      Width           =   2775
   End
   Begin VB.TextBox txtWad2 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1020
      Width           =   2775
   End
   Begin VB.CommandButton cmdDeh2 
      Caption         =   "Select 2nd Deh"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1860
      Width           =   1575
   End
   Begin VB.CommandButton cmdWad2 
      Caption         =   "Select 2nd Wad"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   3180
      TabIndex        =   18
      Top             =   3900
      Width           =   975
   End
   Begin VB.Frame fraIwad 
      Caption         =   "Runs using"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   19
      Top             =   2340
      Width           =   4215
      Begin VB.OptionButton optIwad 
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
         Index           =   7
         Left            =   2700
         TabIndex        =   23
         Top             =   1140
         Width           =   1320
      End
      Begin VB.OptionButton optIwad 
         Caption         =   "Strife"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   5
         Left            =   2700
         TabIndex        =   15
         Top             =   840
         Width           =   1320
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1500
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   13
         Top             =   540
         Width           =   1020
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   1500
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   1
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Width           =   1260
      End
      Begin VB.OptionButton optIwad 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDeh1 
      Caption         =   "Select Deh"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtDeh1 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
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
      Left            =   540
      TabIndex        =   17
      Top             =   3900
      Width           =   975
   End
   Begin VB.CommandButton cmdWad1 
      Caption         =   "Select Wad"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtWad1 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblWadName 
      Caption         =   "Wad Name"
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
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmWadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdClear_Click()

txtName.Text = ""
txtWad1.Text = ""
txtWad2.Text = ""
txtDeh1.Text = ""
txtDeh2.Text = ""
optIwad(1).Value = True

End Sub

Private Sub cmdDeh2_Click()

Call OpenFile("*.deh", "Select DEH Patch", txtDeh2)

End Sub

Private Sub cmdOk_Click()

Select Case WorkMode
    Case "ADD"
    If txtName.Text = "" Then
    MsgBox "You must type a name for this entry", vbExclamation, "Error!"
    ElseIf txtWad1.Text = "" Then
    MsgBox "A PWAD1 path is required", vbExclamation, "Error!"
    Else
    Call SaveData("ADD", txtName, rsWads, "WadName", txtWad1, rsWads, "WadPath", txtWad2, rsWads, "WadPath2", txtDeh1, rsWads, "DehPath", txtDeh2, rsWads, "DehPath2", txtNegro, rsWads, "StartWithIwad", txtFile, rsWads, "WadFile")
    rsWads.MoveFirst
    Unload Me
    End If
    
    Case "EDIT"
    If txtName.Text = "" Then
    MsgBox "You must type a name for this entry", vbExclamation, "Error!"
    ElseIf txtWad1.Text = "" Then
    MsgBox "A PWAD1 path is required", vbExclamation, "Error!"
    Else
    Call SaveData("EDIT", txtName, rsWads, "WadName", txtWad1, rsWads, "WadPath", txtWad2, rsWads, "WadPath2", txtDeh1, rsWads, "DehPath", txtDeh2, rsWads, "DehPath2", txtNegro, rsWads, "StartWithIwad", txtFile, rsWads, "WadFile")
    rsWads.MoveFirst
    Unload Me
    End If
End Select

End Sub

Private Sub cmdDeh1_Click()

Call OpenFile("*.deh", "Select DEH Patch", txtDeh1)

End Sub

Private Sub cmdWad1_Click()

Call OpenFile("*.wad", "Select PWAD", txtWad1)

End Sub


Private Sub cmdWad2_Click()

Call OpenFile("*.wad", "Select PWAD", txtWad2)

End Sub

Private Sub Form_Load()

Select Case WorkMode

Case "ADD": optIwad(0).Value = True
Case "EDIT":
    Call LoadData(txtName, rsWads, "WadName", txtWad1, rsWads, "WadPath", txtWad2, rsWads, "WadPath2", txtDeh1, rsWads, "DehPath", txtDeh2, rsWads, "DehPath2", txtNegro, rsWads, "StartWithIwad", txtFile, rsWads, "WadFile")
    Select Case txtNegro.Text
        Case "Doom": optIwad(0).Value = True
        Case "Doom II": optIwad(1).Value = True
        Case "Plutonia": optIwad(2).Value = True
        Case "Tnt": optIwad(3).Value = True
        Case "Heretic": optIwad(4).Value = True
        Case "Hexen": optIwad(5).Value = True
        Case "Strife": optIwad(6).Value = True
        Case "FreeDoom": optIwad(7).Value = True
    End Select
End Select
End Sub

Private Sub optIwad_Click(Index As Integer)

Select Case Index

Case 0: txtNegro.Text = "Doom"
Case 1: txtNegro.Text = "Doom II"
Case 2: txtNegro.Text = "Plutonia"
Case 3: txtNegro.Text = "Tnt"
Case 4: txtNegro.Text = "Heretic"
Case 5: txtNegro.Text = "Hexen"
Case 6: txtNegro.Text = "Strife"
Case 7: txtNegro.Text = "FreeDoom"

End Select
End Sub
