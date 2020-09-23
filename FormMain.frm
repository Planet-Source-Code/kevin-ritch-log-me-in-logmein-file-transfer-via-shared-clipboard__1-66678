VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "For LogMeIn Users to transfer SMALLER files between Support & Remote PC's (Either Direction)"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FormMain.frx":030A
   ScaleHeight     =   3030
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton HowItWorksButton 
      Height          =   495
      Left            =   9600
      Picture         =   "FormMain.frx":71C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton ReceiveFileButton 
      Caption         =   "RECEIVE FILE"
      Height          =   855
      Left            =   8640
      Picture         =   "FormMain.frx":74CF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton SendFileButton 
      Caption         =   "SEND FILE"
      Height          =   855
      Left            =   240
      Picture         =   "FormMain.frx":7911
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vers. 2.0.1"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUST RUN ON BOTH HOST && CLIENT PC'S AT THE SAME TIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please do remember to vote!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   5385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A program to TRANSFER FILES when using the brilliant ""LogMeIn"" Software"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
 End
End Sub
Private Sub HowItWorksButton_Click()
 M$ = "If you are a LogMeIn user and you need to transfer large files (perhaps in excess of 2 Megs?) then you may wish to consider using a different method such as Zipped email attachments, HTTP upload/download, FTP or even UPGRADE your ''LogMeIn'' Software!"
 M$ = M$ & String$(2, 10) & "Remember, ZIPPING is the best solution for those biggies!  (FREE with Windows XP onwards of course)"
 M$ = M$ & String$(2, 10) & "GOOD NEWS: LogMeIn DOES provide for a SHARED TEXT CLIPBOARD - though it does actually need quite some time to Synchronize between the two PC's."
 M$ = M$ & String$(2, 10) & "What my program (I call it V8Transfer) does, is it translates any file into PURE ASCII and sets the text to the Clipboard."
 M$ = M$ & String$(2, 10) & "The SAVE_AS function merely translates the ASCII TEXT back to BINARY and stuffs it into the filename of your choice!"
 M$ = M$ & String$(2, 10) & "Please do VOTE for this submission if you feel that it is cleverly inventive or resourceful and reckon that this is just ''the bees knees'' :-)"
 M$ = M$ & String$(2, 10) & "THANKS GUYS & DOLLS AT PLANET SOURCE CODE ~ HAVE A REALLY GROOVY DAY!"
 MsgBox M$, vbApplicationModal + vbInformation, "THIS IS HOW IT WORKS Y'ALL!"
End Sub
Private Sub ReceiveFileButton_Click()
 Me.Visible = False
 ReceiveFileForm.Show
End Sub
Private Sub SendFileButton_Click()
 Me.Visible = False
 SendForm.Show
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
 End If
End Sub

