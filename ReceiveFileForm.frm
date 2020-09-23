VERSION 5.00
Begin VB.Form ReceiveFileForm 
   Caption         =   "RECEIVE FILE FORM"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   Icon            =   "ReceiveFileForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton SaveAsButton 
      Caption         =   "SAVE FILE AS..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Larger The File - The Longer You May Have To Wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please do remember to vote!"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Written by Kevin Ritch ( www.GreatCRM.com )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4665
   End
End
Attribute VB_Name = "ReceiveFileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SaveAsButton_Click()
 Dim Pointer As Long
 Dim iSize As Long
 On Error Resume Next
 MkDir "c:\Windows\"
 MkDir "c:\Windows\Temp"
 On Error GoTo Blooper:
 DoEvents
 MsgBox "The larger the file, the longer LogMeIn will take to REFRESH the remote Windows clipboard!"
 FilePickerForm.SaveAsCaution.Visible = True
 FilePickerForm.Show vbModal, Me
 FileName$ = FilePickerForm.File1.Path & IIf(Right$(FilePickerForm.File1, 1) <> "\", "\", "") & FilePickerForm.FileNameText
 FileName$ = Replace$(FileName$, "\\", "\") ' Just say "no" to Network Files!
 Unload FilePickerForm
 If Dir$(FileName$) <> "" Then
  If MsgBox("The file:" & String$(2, 10) & FileName$ & String$(2, 10) & "Already EXISTS!" & String$(2, 10) & "Are you sure that you wish to OVER-WRITE it?", vbApplicationModal + vbDefaultButton2 + vbYesNo + vbQuestion, "OVER-WRITE EXISTING FILE?") = vbNo Then
   FileName$ = ""
  End If
 End If
 If FileName$ = "" Then
  MsgBox "Save_As has been cancelled!"
  Exit Sub
 End If
 Open "c:\Windows\Temp\Temporary.HexFile" For Output As #1
 Print #1, Clipboard.GetText;
 Close
 Open "c:\Windows\Temp\Temporary.HexFile" For Binary Shared As #1
 iSize = LOF(1)
 Open FileName$ For Output As #2
 For Pointer = 1 To (iSize - 1) Step 2
  Seek #1, Pointer
  Print #2, Chr$(Val("&H" & Input$(2, #1)));
 Next Pointer
 Close
 Kill "c:\Windows\Temp\Temporary.HexFile"
 MsgBox "The file:" & String$(2, 10) & FileName$ & String$(2, 10) & "has been created!", vbApplicationModal + vbInformation, "FINISHED - PLEASE CHECK YOUR FILE!"
 Unload Me
 Exit Sub
Blooper:
 MsgBox "There's a problem with saving to that filename!" & String$(2, 10) & FileName$, vbApplicationModal + vbExclamation, "Whoops!"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
 End If
End Sub
Private Sub Form_Load()
 Me.Picture = FormMain.Picture
End Sub
Private Sub Form_Unload(Cancel As Integer)
 FormMain.Visible = True
End Sub
