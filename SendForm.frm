VERSION 5.00
Begin VB.Form SendForm 
   Caption         =   "SEND FILE FORM"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   Icon            =   "SendForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton SendFileButton 
      Caption         =   "PICK FILE TO SEND TO REMOTE"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Try with a SMALL File 1st (Say a small 32k jpg)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4425
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
      TabIndex        =   2
      Top             =   1560
      Width           =   4665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please do remember to vote!"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2025
   End
End
Attribute VB_Name = "SendForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SendFileButton_Click()
 FilePickerForm.SaveAsCaution.Visible = False
 FilePickerForm.Show vbModal, Me
 FileName$ = FilePickerForm.File1.Path & IIf(Right$(FilePickerForm.File1, 1) <> "\", "\", "") & FilePickerForm.FileNameText
 FileName$ = IIf(Dir(FileName$) = "", "", IIf(Right$(FileName$, 1) <> "\", FileName$, ""))
 Unload FilePickerForm
 If FileName$ <> "" Then
  Call ClearClipboard
  Call FileToHex(FileName$)
  Open "c:\Windows\Temp\Temporary.HexFile" For Binary Shared As #1
  Seek #1, 1
  Clipboard.SetText Input$(LOF(1), #1)
  Close
  Kill "c:\Windows\Temp\Temporary.HexFile"
  MsgBox FileName$ & String$(2, 10) & String$(38, "=") & String$(2, 10) & " STORED AS PURE ASCII TEXT ON THE WINDOWS CLIPBOARD" & String$(2, 10) & String$(38, "=") & String$(3, 10) & "On the REMOTE PC, please press the ''RECEIVE FILE'' button!", vbApplicationModal + vbInformation, "FILE STORED TO CLIPBOARD"
  Unload Me
 End If
End Sub
Sub FileToHex(FilenameStr As String)
 Dim i As Long
 On Error Resume Next
 MkDir "c:\Windows\"
 MkDir "c:\Windows\Temp"
 On Error GoTo 0
 Open FilenameStr For Binary Shared As #1
 Open "c:\Windows\Temp\Temporary.HexFile" For Output As #2
 BytesToConvert = LOF(1)
 For i = 1 To BytesToConvert
  Seek #1, i
  a$ = Input$(1, #1)
  Print #2, Right$("00" & Hex$(Asc(a$)), 2);
 Next i
 Close
End Sub
Private Sub Form_Load()
 Me.Picture = FormMain.Picture
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Cancel = True
 Me.Visible = False
 FormMain.Visible = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
 End If
End Sub
