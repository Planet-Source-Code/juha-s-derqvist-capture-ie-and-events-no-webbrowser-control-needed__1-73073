VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents captx As SHDocVw.InternetExplorer
Attribute captx.VB_VarHelpID = -1
Private Sub captx_StatusTextChange(ByVal Text As String)
    Form1.Text1.Text = Text
    Form1.List1.AddItem Now & " - " & Text
End Sub

