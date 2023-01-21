VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entryway"
   ClientHeight    =   3390
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   3060
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13256
            MinWidth        =   2646
            Text            =   "No Wad Selected"
            TextSave        =   "No Wad Selected"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstWads 
      Height          =   3060
      IntegralHeight  =   0   'False
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3135
   End
   Begin VB.Image Snapshot 
      BorderStyle     =   1  'Fixed Single
      Height          =   3045
      Left            =   3180
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPlayUsing 
         Caption         =   "Play Using"
         Begin VB.Menu mnuGZDoom 
            Caption         =   "GZDoom"
         End
         Begin VB.Menu mnuZDoom 
            Caption         =   "ZDoom"
         End
         Begin VB.Menu mnuChocolate 
            Caption         =   "Chocolate Doom"
         End
         Begin VB.Menu mnuEdge 
            Caption         =   "EDGE"
         End
         Begin VB.Menu mnuSkulltag 
            Caption         =   "Skulltag"
         End
         Begin VB.Menu mnuPrboom 
            Caption         =   "prBoom"
         End
         Begin VB.Menu mnuEternity 
            Caption         =   "Eternity Engine"
         End
         Begin VB.Menu mnuVavoom 
            Caption         =   "Vavoom"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuManageWads 
      Caption         =   "Manage Wads"
      Begin VB.Menu mnuAddEntry 
         Caption         =   "Add New Entry"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditCurrentEntry 
         Caption         =   "Edit Entry"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Entry"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSelectExe 
         Caption         =   "IWAD/EXE Options"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SnapFile As String
Dim WadFile As String



Private Sub Form_Activate()


Call CheckFieldNullity(rsIwads, "IwadPath")
Call CheckFieldNullity(rsWads, "WadPath")
Call CheckFieldNullity(rsWads, "WadName")
Call CheckFieldNullity(rsWads, "DehPath")
Call CheckFieldNullity(rsWads, "WadPath2")
Call CheckFieldNullity(rsWads, "DehPath2")
Call CheckFieldNullity(rsWads, "WadFile")
Call CheckFieldNullity(rsExe, "ExePath")
Call QueProgramaDeMierdaEste

rsWads.Requery
rsExe.Requery
rsIwads.Requery
lstWads.Clear
 
    If Not rsWads.EOF Then
        rsWads.MoveFirst
        Do While Not rsWads.EOF
            lstWads.AddItem rsWads!wadname
            lstWads.ItemData(lstWads.NewIndex) = rsWads!WadNumber
            rsWads.MoveNext
        Loop
        lstWads.ListIndex = 0
    Else
        lstWads.ListIndex = -1
    End If
   
rsExe.FindFirst "ExeName Like 'GZDoom'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuGZDoom.Enabled = True
Else
mnuGZDoom.Enabled = False
End If

rsExe.FindFirst "ExeName Like 'Zdoom'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuZDoom.Enabled = True
Else
mnuZDoom.Enabled = False

End If

rsExe.FindFirst "ExeName Like 'Chocolate & Doom'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuChocolate.Enabled = True
Else
mnuChocolate.Enabled = False

End If
   
rsExe.FindFirst "ExeName Like 'EDGE'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuEdge.Enabled = True
Else
mnuEdge.Enabled = False
End If
   
rsExe.FindFirst "ExeName Like 'Skulltag'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuSkulltag.Enabled = True
Else
mnuSkulltag.Enabled = False
End If
   
rsExe.FindFirst "ExeName Like 'prBoom'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuPrboom.Enabled = True
Else
mnuPrboom.Enabled = False
End If
   
rsExe.FindFirst "ExeName Like 'Eternity'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuEternity.Enabled = True
Else
mnuEternity.Enabled = False
End If
   
rsExe.FindFirst "ExeName Like 'Vavoom'"
If Dir(rsExe!ExePath) <> "" = True Then
mnuVavoom.Enabled = True
Else
mnuVavoom.Enabled = False

End If
   
rsExe.MoveFirst
   
End Sub

Private Sub Form_Load()

Call EstablishingBattleControlStandBy

End Sub

Private Sub lstWads_Click()
rsIwads.Requery
rsWads.Requery
rsWads.FindFirst "WadNumber=" & Str(lstWads.ItemData(lstWads.ListIndex))
rsIwads.FindFirst "IwadName Like'*" & rsWads!StartWithIwad & "*'"
SnapFile = Left(rsWads!WadFile, Len(rsWads!WadFile) - 4)
    If Dir("Snaps\" & SnapFile & ".jpg") <> "" Then
        Snapshot.Picture = LoadPicture("Snaps\" & SnapFile & ".jpg")
    Else
        Snapshot.Picture = LoadPicture("Snaps\SnapNotAvailable.jpg")
    End If
    
StatusBar1.Panels(1).Text = "Base Game: " & rsWads!StartWithIwad
    
End Sub

Private Sub mnuEditWadList_Click()

Call WadForm("EDIT")

End Sub

Private Sub mnuAddEntry_Click()

Call WadForm("ADD")

End Sub

Private Sub mnuDelete_Click()

    If lstWads.ListIndex > -1 Then
        rsWads.Delete
        lstWads.RemoveItem lstWads.ListIndex
        rsWads.MoveLast
    Else
        lstWads.ListIndex = -1
    End If

End Sub

Private Sub mnuEditCurrentEntry_Click()

    If lstWads.ListIndex = -1 Then
        lstWads.ListIndex = -1
    Else
        Call WadForm("EDIT")
    End If

End Sub

Private Sub mnuEternity_Click()

rsExe.FindFirst "ExeNumber=7"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuExit_Click()

Unload Me

End Sub



Private Sub mnuChocolate_Click()

rsExe.FindFirst "ExeNumber=3"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuGZDoom_Click()

rsExe.FindFirst "ExeNumber=1"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuEdge_Click()

rsExe.FindFirst "ExeNumber=4"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuSkulltag_Click()

rsExe.FindFirst "ExeNumber=5"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuZDoom_Click()

rsExe.FindFirst "ExeNumber=2"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuPrboom_Click()

rsExe.FindFirst "ExeNumber=6"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

Private Sub mnuSelectExe_Click()

frmOptions.Show vbModal

End Sub

Private Sub mnuVavoom_Click()

rsExe.FindFirst "ExeNumber=8"
Shell Chr(34) & rsExe!ExePath & Chr(34) & " -file " & Chr(34) & rsWads!WadPath & Chr(34) & " -iwad " & Chr(34) & rsIwads!IwadPath & Chr(34), vbNormalFocus

End Sub

