VERSION 5.00
Begin VB.Form frmVB6toVB5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB6 to VB5"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVB6toVB5.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Text            =   "Retained=0"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3555
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4035
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "frmVB6toVB5.frx":0442
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3045
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "coded by: wizdum"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Thanx to ----Mouse---- of tbh for teaching me about dragging and dropping files and have fun with the source!"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmVB6toVB5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'hope u have fun with this source, wiz
Const ChunkSize = 4096 '4096 works the best 4 me

Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim file
If data.GetFormat(vbCFFiles) Then
For Each file In data.Files
If UCase(Right(file, 4)) = UCase(".vbp") Then
Text4.Text = file
vb6tovb5 Text4.Text, Text2.Text, Text3.Text
Else
MsgBox "Make sure the file extension is .vbp!", vbCritical, "Error"
End If
Next file
Else
MsgBox "Other Error, Please make sure it is a vb Project!!!", vbCritical, "Error"
End If
End Sub

Public Sub vb6tovb5(FName$, IDString$, NString$)
Dim tempstring, Msg
Dim PosString, WhereString
Dim FileNumber, A$, NewString$
Dim AString As String * ChunkSize
Dim IsChanged As Boolean
Dim BlockIsChanged As Boolean
Dim NumChanges As Integer
IsChanged = False
BlockIsChanged = False
On Error Resume Next
FileNumber = FreeFile
PosString = 1
WhereString = 0
AString = Space$(ChunkSize)
'Make sure strings have same size
If Len(IDString$) > Len(NString$) Then
NewString$ = NString$ + Space$(Len(IDString$) - Len(NString$))
Else
NewString$ = Left$(NString$, Len(IDString$))
End If
Open FName$ For Binary As FileNumber
NumChanges = 0
If LOF(FileNumber) < ChunkSize Then
A$ = Space$(LOF(FileNumber))
Get #FileNumber, 1, A$
WhereString = FindInString(1, A$, IDString$)
Else
A$ = Space$(ChunkSize)
Get #FileNumber, 1, A$
WhereString = FindInString(1, A$, IDString$)
End If
Do
While WhereString <> 0
tempstring = Left$(A$, WhereString - 1) & NewString$ & Mid$(A$, WhereString + Len(NewString$))
A$ = tempstring
NumChanges = NumChanges + 1
IsChanged = True
BlockIsChanged = True
WhereString = FindInString(WhereString + 1, A$, IDString$)
Wend
If BlockIsChanged Then
Put #FileNumber, PosString, A$
BlockIsChanged = False
End If
PosString = ChunkSize + PosString - Len(IDString$)
' If we're finished, exit the loop
If EOF(FileNumber) Or PosString > LOF(FileNumber) Then
Exit Do
End If
' Get the next chunk to scan
If PosString + ChunkSize > LOF(FileNumber) Then
A$ = Space$(LOF(FileNumber) - PosString + 1)
Get #FileNumber, PosString, A$
WhereString = FindInString(1, A$, IDString$)
Else
A$ = Space$(ChunkSize)
Get #FileNumber, PosString, A$
WhereString = FindInString(1, A$, IDString$)
End If
Loop Until EOF(FileNumber) Or PosString > LOF(FileNumber)
If IsChanged = True Then
MsgBox FName$ & " is now compatible with VB5", vbInformation, "File converted to VB5"
Else
MsgBox "File has not been converted to VB5", vbCritical, "Could not convert"
End If
End Sub

Private Function FindInString(StartPos As Integer, StrToSearch As String, _
StrToFind As String) As Integer
If Check1.Value = 0 Then
FindInString = InStr(StartPos, UCase(StrToSearch), UCase(StrToFind))
Else
FindInString = InStr(StartPos, StrToSearch, StrToFind)
End If
End Function

