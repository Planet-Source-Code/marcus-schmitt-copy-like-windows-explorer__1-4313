VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Copy like Explorer"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Copy C:\temp to C:\temptest"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copy files like Explorer
'
' M.Schmitt
' Marcus-Schmitt@gmx.de
' www.marcus.schmitt.notrix.de

Private Sub Command1_Click()
    'CopyFile
    'Source, Destination, Ask to overwrite if files already exists, Visible Window
    x = CopyFile("C:\temp", "C:\tempneu", True, True)
    If Not x Then MsgBox "Error !", vbCritical, "Oops"
End Sub
