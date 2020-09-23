VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Capture IE - reference direct no webbrowser control needed"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   7275
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   11655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   495
      Left            =   10920
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "juhax.com"
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ie ref
Dim captIE As New SHDocVw.InternetExplorer
' capture class ie
Dim captieclass As New Class1

Private Sub Command1_Click()
captIE.Visible = True
captIE.Navigate "www.juhax.com"
End Sub

Private Sub Form_Load()
Set captieclass.captx = captIE
End Sub

