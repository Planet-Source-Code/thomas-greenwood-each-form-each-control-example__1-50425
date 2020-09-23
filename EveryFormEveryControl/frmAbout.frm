VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OKButton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label emaillbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "thomasgreenwood@2die4.com"
      Height          =   195
      Left            =   1095
      TabIndex        =   3
      Top             =   840
      Width           =   2490
   End
   Begin VB.Label titlelbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Every Form, Every Control Example"
      Height          =   195
      Left            =   1095
      TabIndex        =   2
      Top             =   120
      Width           =   2490
   End
   Begin VB.Label namelbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made by Thomas Greenwood"
      Height          =   195
      Left            =   1095
      TabIndex        =   0
      Top             =   480
      Width           =   2490
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OKButton_Click()
Unload Me
End Sub
