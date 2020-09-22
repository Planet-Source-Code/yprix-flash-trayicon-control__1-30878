VERSION 5.00
Object = "{910882D4-E694-11D5-A104-00064F006EED}#14.0#0"; "TrayControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flash Tray Control Demo Application"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1725
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2115
      Width           =   5100
   End
   Begin TrayControl.TrayIcon TrayIcon1 
      Left            =   4725
      Top             =   60
      _ExtentX        =   1058
      _ExtentY        =   1058
      Icon            =   "MainForm.frx":030A
      ToolTip         =   "Email checker"
      Caption         =   "Flash TrayIcon"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "For bug reports or comments contact me @ cyprix@email.ro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   885
      TabIndex        =   1
      Top             =   1395
      Width           =   3240
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "TO DO: put your stuff here:"
      Height          =   255
      Left            =   1350
      TabIndex        =   0
      Top             =   720
      Width           =   2025
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    With TrayIcon1
        .MenuItems.Add mtMenuItem, "About FlashTrayIcon Control"
        .MenuItems.Add mtMenuItem, "Help"
        .MenuItems.Add mtSeparator
        .MenuItems.Add mtMenuItem, "Disabled"
        .MenuItems.Add mtMenuItem, "Check"
        .MenuItems.Add mtSeparator
        .MenuItems.Add mtMenuItem, "Hide form"
        .MenuItems.Add mtMenuItem, "Show form"
        .MenuItems.Add mtSeparator
        .MenuItems.Add mtMenuItem, "Exit"
        .MenuItems(8).Grayed = True
        .MenuItems(4).Grayed = True
        .MenuItems(2).MenuItems.Add mtMenuItem, "Control documentation"
        .MenuItems(2).MenuItems.Add mtMenuItem, "Online reference"
        .MenuItems(2).MenuItems.Add mtMenuItem, "FAQ's"
        Set .MenuItems(10).Icon = ImageList1.ListImages(1).Picture
        .AtachToSysTray
        
    End With
    
End Sub








Private Sub TrayIcon1_OnMenuItemSelect(ByVal ItemIndex As Long)
   
    Text1.Text = Text1.Text + CStr(ItemIndex) + vbNewLine
    
    If ItemIndex = 1010000000 Then
        TrayIcon1.AboutBox
    End If
    
    If ItemIndex = 1020100000 Then
        App.HelpFile = App.Path & "\TrayControl.chm"
        SendKeys "{F1}"
        Me.Show
    End If
    
    If ItemIndex = 1020200000 Then
        MsgBox "Go to http://cyprix.topcities.com for that!!!"
    End If
    
    If ItemIndex = 1020300000 Then
        MsgBox "Go to http://cyprix.topcities.com for that!!!"
    End If
    
    If ItemIndex = 1050000000 Then
        TrayIcon1.MenuItems(5).Checked = Not TrayIcon1.MenuItems(5).Checked
        If TrayIcon1.MenuItems(5).Caption = "Uncheck" Then
            TrayIcon1.MenuItems(5).Caption = "Check"
        Else
            TrayIcon1.MenuItems(5).Caption = "Uncheck"
        End If
        TrayIcon1.Update
    End If
    
    If ItemIndex = 1070000000 Then
        Me.WindowState = 1
        Me.Hide
        TrayIcon1.MenuItems(8).Grayed = False
        TrayIcon1.MenuItems(7).Grayed = True
        TrayIcon1.Update
    End If
    
    If ItemIndex = 1080000000 Then
        Me.Show
        Me.WindowState = 0
        TrayIcon1.MenuItems.Item(8).Grayed = True
        TrayIcon1.MenuItems.Item(7).Grayed = False
        TrayIcon1.Update
    End If
    
    If ItemIndex = 1100000000 Then
        TrayIcon1.DetachFromSysTray
        Unload Me
    End If
    


End Sub

