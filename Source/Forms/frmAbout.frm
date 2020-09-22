VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'None
   Caption         =   "  About"
   ClientHeight    =   3195
   ClientLeft      =   3405
   ClientTop       =   2700
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   165
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1815
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   345
      Width           =   4215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3585
      TabIndex        =   0
      Top             =   2685
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   3180
      Left            =   0
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Image1_Click()

End Sub


