VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   653
      TabIndex        =   2
      Text            =   "Press Button to set Forms Captions to this text"
      Top             =   743
      Width           =   3375
   End
   Begin VB.CommandButton capButton 
      Caption         =   "Example Button"
      Height          =   855
      Left            =   1373
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1463
      Width           =   2055
   End
   Begin VB.Label exlbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Example Label"
      Height          =   195
      Left            =   1733
      TabIndex        =   0
      Top             =   383
      Width           =   1035
   End
   Begin VB.Menu AboutClick 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AboutClick_Click()
SetForms
frmAbout.Show 1
End Sub


Private Sub capButton_Click()
SetTitles Text1.Text
End Sub


