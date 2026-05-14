VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Headerbox 
   Caption         =   "Student Details"
   ClientHeight    =   4540
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   7040
   OleObjectBlob   =   "Headerbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Headerbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub UserForm_initialize()

Cmbcourse.AddItem "BCA"
Cmbcourse.AddItem "BSc IT"
Cmbcourse.AddItem "BSc Cs"
Cmbcourse.AddItem "BCom"

Lstskills.AddItem "Excel"
Lstskills.AddItem "PowerBI"
Lstskills.AddItem "VBA"

srlage.Min = 18
srlage.Max = 50

tglinactive.Caption = "Inactive"


End Sub

Private Sub srlage_Change()

txtage.Value = srlage.Value

End Sub

Private Sub tglinactive_Click()

If tglinactive.Value = True Then
    tglinactive.Caption = "Active"
Else
    tglinactive.Caption = "Inactive"
End If

End Sub
Private Sub cmdsubmit_Click()

Dim hostel As String
Dim gender As String
Dim lastrow As Long

'Hotel value
hostel = IIf(chkhostel.Value, "Yes", "No")

'Gender validation
If optmale.Value = True Then
    gender = "Male"
ElseIf optfemale.Value = True Then
    gender = "Female"
End If

'Name validation
If txtname.Value = "" Then
MsgBox "Enter Name", vbExclamation
Exit Sub
End If

'output
lastrow = Sheets("student_record").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Sheets("student_record").Cells(lastrow, 1).Value = txtname.Value
    Sheets("student_record").Cells(lastrow, 2).Value = txtage.Value
    Sheets("student_record").Cells(lastrow, 3).Value = Cmbcourse.Value
    Sheets("student_record").Cells(lastrow, 4).Value = Lstskills.Value
    Sheets("student_record").Cells(lastrow, 5).Value = hostel
    Sheets("student_record").Cells(lastrow, 6).Value = gender
    Sheets("student_record").Cells(lastrow, 7).Value = tglinactive.Caption
    
MsgBox "Data added successfully"

End Sub
Private Sub cmdclear_Click()

    'Clear text box
    txtname.Value = ""
    txtage.Value = ""
    
    'Clear combo box
    Cmbcourse.Value = ""
    
    'Clear list box
    Lstskills.Value = ""
    
    'Clear checkbox
    chkhostel.Value = False
    
    'Clear option box
    optmale.Value = False
    optfemale.Value = False
    
    'Reset toggle button
    tglinactive.Value = False
    tglinactive.Caption = "Active"

End Sub

Private Sub UserForm_Activate()

Dim username As String

username = InputBox("Enter your name")

If username <> "" Then
    lblheading.Caption = "Welcome" & "" & username
End If

End Sub

