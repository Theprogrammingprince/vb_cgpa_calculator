VERSION 5.00
Begin VB.Form Group 
   Caption         =   "Group E CGPA CALCULATOR"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17625
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   17625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGpa 
      Height          =   735
      Left            =   4200
      TabIndex        =   7
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton btnCalculate 
      BackColor       =   &H80000018&
      Caption         =   "CALCULATE CGPA"
      Height          =   1335
      Left            =   4680
      MaskColor       =   &H0000FFFF&
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton btnRecord 
      Caption         =   "RECORD THE COURSE"
      Height          =   1215
      Left            =   4800
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtCredit 
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox cboGrade 
      Height          =   315
      ItemData        =   "CGPA_CALCULATOR_IN_VB.frx":0000
      Left            =   4560
      List            =   "CGPA_CALCULATOR_IN_VB.frx":0002
      TabIndex        =   1
      Text            =   "SELECT GRADE"
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "GPA"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "CREDIT UNIT"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "GRADE"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private grades() As String
Private credits() As Double
Private courseCount As Long
Private Sub Command1_Click()

End Sub

Private Sub btnCalculate_Click()
    If courseCount = 0 Then
        MsgBox "No courses recorded.", vbExclamation
        Exit Sub
    End If
    
    Dim totalGradePoints As Double
    Dim totalCredits As Double
    Dim i As Long
    
    totalGradePoints = 0
    totalCredits = 0
    
    ' Calculate grade points
    For i = 0 To courseCount - 1
        Dim gp As Double
        Select Case grades(i)
            Case "A": gp = 5
            Case "B": gp = 4
            Case "C": gp = 3
            Case "D": gp = 2
            Case "E": gp = 1
            Case "F": gp = 0
        End Select
        totalGradePoints = totalGradePoints + (gp * credits(i))
        totalCredits = totalCredits + credits(i)
    Next i
    
    ' Compute CGPA
    If totalCredits > 0 Then
        Dim cgpa As Double
        cgpa = totalGradePoints / totalCredits
        txtGpa.Text = Format(cgpa, "0.00")
    Else
        MsgBox "Total credit hours cannot be zero.", vbExclamation
    End If
End Sub

Private Sub btnRecord_Click()
    ' Check if grade is selected
    If cboGrade.ListIndex = -1 Then
        MsgBox "Please select a grade.", vbExclamation
        Exit Sub
    End If
    
    ' Validate credit hours
    If Not IsNumeric(txtCredit.Text) Then
        MsgBox "Please enter a valid number for credit hours.", vbExclamation
        txtCredit.SetFocus
        Exit Sub
    End If
    
    Dim credit As Double
    credit = CDbl(txtCredit.Text)
    If credit <= 0 Then
        MsgBox "Credit hours must be positive.", vbExclamation
        txtCredit.SetFocus
        Exit Sub
    End If
    
    ' Record the course
    ReDim Preserve grades(courseCount)
    ReDim Preserve credits(courseCount)
    grades(courseCount) = cboGrade.List(cboGrade.ListIndex)
    credits(courseCount) = credit
    courseCount = courseCount + 1
    
    ' Clear inputs
    cboGrade.ListIndex = -1
    txtCredit.Text = ""
    
    ' Notify user (optional)
    MsgBox "Course recorded successfully!", vbInformation
End Sub

Private Sub Form_Load()
    cboGrade.AddItem "A"
    cboGrade.AddItem "B"
    cboGrade.AddItem "C"
    cboGrade.AddItem "D"
    cboGrade.AddItem "E"
    cboGrade.AddItem "F"
    courseCount = 0
    ReDim grades(0)
    ReDim credits(0)
End Sub

