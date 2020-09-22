VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hello world"
   ClientHeight    =   420
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11220
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      Picture         =   "frmMenu.frx":014A
      ScaleHeight     =   675
      ScaleWidth      =   1020
      TabIndex        =   0
      ToolTipText     =   "Net Watch"
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prvIconTray As Long
Public prvParent As TrayIcon
Public prvCreated As Boolean







Private Sub Picture1_Click()

    prvParent.NotifyRightClick
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     X = X / Screen.TwipsPerPixelX
     Select Case X
        Case WM_LBUTTONDOWN
              MenuTrack Me
              prvParent.NotifyLeftClick
        Case WM_RBUTTONDOWN
              MenuTrack Me
              prvParent.NotifyRightClick
        Case WM_MOUSEMOVE
              prvParent.NotifyMouseMove
        Case WM_LBUTTONDBLCLK
              prvParent.NotifyDBLClick
    End Select

End Sub


Public Sub CreateIcon()
    
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hWnd = Picture1.hWnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = Trim$(prvParent.ToolTip) + Chr(0)
    prvIconTray = Shell_NotifyIcon(NIM_ADD, Tic)
    Hook Me, prvParent
    prvCreated = True
    
End Sub
Public Sub DeleteIcon()
    
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hWnd = Picture1.hWnd
    Tic.uID = 1&
    prvIconTray = Shell_NotifyIcon(NIM_DELETE, Tic)
    UnHook
    prvCreated = False
    
End Sub
Public Sub UpdateIcon()
        
        Dim Tic As NOTIFYICONDATA
        Tic.cbSize = Len(Tic)
        Tic.hWnd = Picture1.hWnd
        Tic.uID = 1&
        Tic.uFlags = NIF_DOALL
        Tic.uCallbackMessage = WM_MOUSEMOVE
        Tic.hIcon = Picture1.Picture
        Tic.szTip = Trim$(prvParent.ToolTip) + Chr(0)
        prvIconTray = Shell_NotifyIcon(NIM_MODIFY, Tic)
        UnHook
        Hook Me, prvParent
        
End Sub

