VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub del_reg_sub(inkey As String)
On Error Resume Next
Dim spreg() As String
spreg = Split(inkey, "\")
Select Case UCase(spreg(0))
Case "HKEY_CLASSES_ROOT"
inkeyL = &H80000000
Case "HKEY_CURRENT_USER"
inkeyL = &H80000001
Case "HKEY_LOCAL_MACHINE"
inkeyL = &H80000002
Case "HKEY_USERS"
inkeyL = &H80000003
End Select
Dim sectionl As Integer
sectionl = Len(inkey) - Len(spreg(0))
section = Mid(inkey, Len(inkey) - sectionl + 2)
Dim reg As New cRegistry
Dim Key() As String
Dim sec() As String
Dim Value() As String
Dim v As Long
Dim k As Long
Dim s As Long
reg.ClassKey = inkeyL
reg.SectionKey = section

10: reg.EnumerateSections sec, s
reg.EnumerateValues Value, v
For j = 1 To v
reg.ValueKey = Value(j)
reg.DeleteValue
Next
While s > 0
For i = 1 To s
'reg.EnumerateValues value, v
del_reg_sub inkey & "\" & sec(i)
GoTo 10
Next
Exit Sub
Wend
reg.DeleteKey
End Sub

Private Sub Command1_Click()
'del_reg_sub "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft"
End Sub

