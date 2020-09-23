VERSION 5.00
Begin VB.Form FilePickerForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Basic File Picker "
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "FilePickerForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton PickedButton 
      Caption         =   "ACCEPT"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Timer InitializeTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   3240
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox FileNameText 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   9015
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   9015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin VB.Label SaveAsCaution 
      Caption         =   $"FilePickerForm.frx":0442
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   6015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FilePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Picked As Boolean
Private Sub CancelButton_Click()
 Picked = False
 Unload Me
End Sub
Private Sub File1_DblClick()
 Call PickedButton_Click
End Sub
Private Sub PickedButton_Click()
 Picked = True
 Unload Me
End Sub
Private Sub Dir1_Click()
 File1.Path = Dir1.List(Dir1.ListIndex)
 On Error Resume Next
 File1.ListIndex = 0
End Sub
Private Sub Drive1_Change()
 On Error Resume Next
 Dir1.Path = Left$(Drive1, 2)
 Call Dir1_Click
End Sub
Private Sub File1_Click()
 FileNameText = File1
End Sub
Private Sub FileNameText_GotFocus()
 FileNameText.SelStart = 0
 FileNameText.SelLength = Len(FileNameText)
End Sub
Private Sub Form_Load()
 InitializeTimer.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
 FileNameText = IIf(Picked, FileNameText, "")
 If Me.Visible Then
  Me.Visible = False
  Cancel = True
 Else
  Cancel = False
 End If
End Sub
Private Sub InitializeTimer_Timer()
 InitializeTimer.Enabled = False
 Call Dir1_Click
 Call File1_Click
End Sub
