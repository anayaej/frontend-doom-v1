Attribute VB_Name = "Module1"
Option Explicit

Global OkClick As Boolean

Public dbWads As Database
Public rsWads As Recordset
Public rsExe As Recordset
Public rsIwads As Recordset

Public WorkMode


Sub EstablishingBattleControlStandBy()

Set dbWads = OpenDatabase("database.mdb")
Set rsExe = dbWads.OpenRecordset("Exe", dbOpenDynaset)
Set rsWads = dbWads.OpenRecordset("Wads", dbOpenDynaset)
Set rsIwads = dbWads.OpenRecordset("Iwads", dbOpenDynaset)

End Sub

Sub WadForm(ByVal iWorkMode As String)

Select Case iWorkMode
    Case "ADD"
        WorkMode = "ADD"
        frmWadForm.Caption = "Add New Entry"
    Case "EDIT"
        WorkMode = "EDIT"
        frmWadForm.Caption = "Edit " & rsWads!wadname
End Select

frmWadForm.Show vbModal

End Sub

Sub OpenFile(ByVal Pattern As String, ByVal Caption As String, ByVal Destination As Object)
OkClick = False

frmOpen.FileSelect.Pattern = Pattern
frmOpen.Caption = Caption
frmOpen.Show vbModal

Do
DoEvents
Loop Until OkClick = True

frmWadForm.txtFile.Text = frmOpen.FileSelect.FileName
    If Len(frmOpen.FileSelect.Path) = 3 Then
        Destination = frmOpen.FileSelect.Path & frmOpen.FileSelect.FileName
    Else
        Destination = frmOpen.FileSelect.Path & "\" & frmOpen.FileSelect.FileName
    End If

OkClick = False
Unload frmOpen

End Sub

Function CheckExeistance(ByVal ExeNumber As Integer) As Long

rsExe.FindFirst "ExeNumber=" & ExeNumber
    
    If rsExe.NoMatch = False And rsExe!Enabled = 1 Then
        CheckExeistance = 1
    Else
        CheckExeistance = 0
    End If

End Function
          


Sub SaveData(ByVal Mode As String, ByVal Source As Object, ByVal Recordset As Recordset, ByVal Field As String, Optional ByVal Source2 As Object, Optional ByVal Recordset2 As Recordset, Optional ByVal Field2 As String, Optional ByVal Source3 As Object, Optional ByVal Recordset3 As Recordset, Optional ByVal Field3 As String, Optional ByVal Source4 As Object, Optional ByVal Recordset4 As Recordset, Optional ByVal Field4 As String, Optional ByVal Source5 As Object, Optional ByVal Recordset5 As Recordset, Optional ByVal Field5 As String, Optional ByVal Source6 As Object, Optional ByVal Recordset6 As Recordset, Optional ByVal Field6 As String, Optional ByVal Source7 As Object, Optional ByVal Recordset7 As Recordset, Optional ByVal Field7 As String)

On Error Resume Next

    Select Case Mode
        Case "ADD": Recordset.AddNew
        Case "EDIT": Recordset.Edit
    End Select

Recordset.Fields(Field).Value = Source

    If IsMissing(Recordset2) = False And IsMissing(Field2) = False And IsMissing(Source2) = False Then Recordset2.Fields(Field2).Value = Source2
    If IsMissing(Recordset3) = False And IsMissing(Field3) = False And IsMissing(Source3) = False Then Recordset3.Fields(Field3).Value = Source3
    If IsMissing(Recordset4) = False And IsMissing(Field4) = False And IsMissing(Source4) = False Then Recordset4.Fields(Field4).Value = Source4
    If IsMissing(Recordset5) = False And IsMissing(Field5) = False And IsMissing(Source5) = False Then Recordset5.Fields(Field5).Value = Source5
    If IsMissing(Recordset6) = False And IsMissing(Field6) = False And IsMissing(Source6) = False Then Recordset6.Fields(Field6).Value = Source6
    If IsMissing(Recordset7) = False And IsMissing(Field7) = False And IsMissing(Source7) = False Then Recordset7.Fields(Field7).Value = Source7

Recordset.Update

End Sub

Sub LoadData(ByVal Destination As Object, ByVal Recordset As Recordset, ByVal Field As String, Optional ByVal Destination2 As Object, Optional ByVal Recordset2 As Recordset, Optional ByVal Field2 As String, Optional ByVal Destination3 As Object, Optional ByVal Recordset3 As Recordset, Optional ByVal Field3 As String, Optional ByVal Destination4 As Object, Optional ByVal Recordset4 As Recordset, Optional ByVal Field4 As String, Optional ByVal Destination5 As Object, Optional ByVal Recordset5 As Recordset, Optional ByVal Field5 As String, Optional ByVal Destination6 As Object, Optional ByVal Recordset6 As Recordset, Optional ByVal Field6 As String, Optional ByVal Destination7 As Object, Optional ByVal Recordset7 As Recordset, Optional ByVal Field7 As String)
On Error Resume Next

Destination = Recordset.Fields(Field).Value

If IsMissing(Recordset2) = False And IsMissing(Field2) = False And IsMissing(Destination2) = False Then Destination2 = Recordset2.Fields(Field2).Value
If IsMissing(Recordset3) = False And IsMissing(Field3) = False And IsMissing(Destination3) = False Then Destination3 = Recordset2.Fields(Field3).Value
If IsMissing(Recordset4) = False And IsMissing(Field4) = False And IsMissing(Destination4) = False Then Destination4 = Recordset2.Fields(Field4).Value
If IsMissing(Recordset5) = False And IsMissing(Field5) = False And IsMissing(Destination5) = False Then Destination5 = Recordset2.Fields(Field5).Value
If IsMissing(Recordset6) = False And IsMissing(Field6) = False And IsMissing(Destination6) = False Then Destination6 = Recordset2.Fields(Field6).Value
If IsMissing(Recordset7) = False And IsMissing(Field7) = False And IsMissing(Destination7) = False Then Destination7 = Recordset2.Fields(Field7).Value
   
End Sub

Sub CheckFieldNullity(ByVal Recordset As Recordset, ByVal Field As String)

If Recordset.EOF Then
    Recordset.AddNew
    Recordset.Fields(Field).Value = ""
    Recordset.Update
Else

Recordset.MoveFirst
    
    Do While Not Recordset.EOF
        If IsNull(Recordset.Fields(Field).Value) Then
            Recordset.Edit
            Recordset.Fields(Field).Value = ""
            Recordset.Update
            Recordset.MoveNext
        Else
            Recordset.MoveNext
        End If
    Loop
End If
Recordset.MoveFirst
End Sub

Sub QueProgramaDeMierdaEste()

rsWads.MoveFirst
    Do While Not rsWads.EOF
        If Dir(rsWads!WadPath) <> "" = True Then
            rsWads.MoveNext
        Else
            rsWads.Delete
            rsWads.MoveNext
        End If
    Loop
rsWads.MoveFirst
End Sub
