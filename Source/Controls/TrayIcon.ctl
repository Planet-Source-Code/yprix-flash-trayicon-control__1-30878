VERSION 5.00
Begin VB.UserControl TrayIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "TrayIcon.ctx":0000
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   41
   ToolboxBitmap   =   "TrayIcon.ctx":000F
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   -15
      Picture         =   "TrayIcon.ctx":0321
      Stretch         =   -1  'True
      Top             =   30
      Width           =   600
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
' Name: Flash System TrayIcon
' Purpose: Active-X Component to create system tray enabled applications
' Functions:
' Properties:
' Methods:
' Author: Ciprian Sorlea
' Start: 15.01.2002
' Modified: 15.01.2002
'===============================================================================
Option Explicit

Private prvIcon As Picture
Private prvCaptionPicture As Picture
Private prvToolTip As String
Private prvCaption As String
Private prvMenuForm As frmMenu
Private prvMenuItems As MenuItems
'=============================
' Name:Eveniment OnMouseMove
' Purpose: User is moving mouse over the system tray icon
' Remarks:
' Author:
'=============================
Event OnMouseMove()
Attribute OnMouseMove.VB_Description = "Event fired when user moves the mouse cursor over the system tray icon"

'=============================
' Name:Eveniment OnLeftClick
' Purpose: User has pressed the left button on the system tray icon
' Remarks:
' Author:
'=============================
Event OnLeftClick()
Attribute OnLeftClick.VB_Description = "Event fired when user leftclicks over the system tray icon"

'=============================
' Name:Eveniment OnRightClick
' Purpose: User has pressed the right button on the system tray icon
' Remarks:
' Author:
'=============================
Event OnRightClick()
Attribute OnRightClick.VB_Description = "Event fired when userrightclicks over the system tray icon"

'=============================
' Name:Eveniment OnDblClick
' Purpose: User has double clicked the system tray icon
' Remarks:
' Author:
'=============================
Event OnDblClick()
Attribute OnDblClick.VB_Description = "Event fired when user doubleclicks the system tray icon"

'=============================
' Name:Eveniment OnMenuItemSelect
' Input:
'   ByVal ItemIndex as Long - The index of the selected menu item (except separators)
' Output:
' Purpose: User has selected an option in the system tray menu
' Remarks:
' Author:
'=============================
Event OnMenuItemSelect(ByVal ItemIndex As Long)

'===============================================================================
' Name: Property Get Icon
' Input:
'
' Output:
'    Picture - The icon in the system tray
' Purpose: Gets the icon
' Remarks:
'===============================================================================
'MemberInfo=11,0,0,0
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Sets/gets the system tray icon"
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Icon = prvIcon
End Property
'===============================================================================
' Name: Property Set Icon
' Input:
'    ByVal New_Image as Picture - The icon in the system tray
' Output:
' Purpose: Sets the icon
' Remarks: Only icon types allowed!!!
'===============================================================================
Public Property Set Icon(ByVal New_Image As Picture)
    Set prvIcon = New_Image
    Set prvMenuForm.Picture1.Picture = prvIcon
    PropertyChanged "Icon"
    prvMenuForm.UpdateIcon
End Property

'===============================================================================
' Name: Property Get CaptionPicture
' Input:
'
' Output:
'    Picture - The picture contained on system tray menubar. If missing then the default gradient caption is painted
' Purpose: Gets the caption picture
' Remarks:
'===============================================================================
'MemberInfo=11,0,0,0
Public Property Get CaptionPicture() As Picture
    Set CaptionPicture = prvCaptionPicture
End Property
'===============================================================================
' Name: Property Set CaptionPicture
' Input:
'    ByVal New_Image as Picture - The picture contained on system tray menubar. If missing then the default gradient caption is painted
' Output:
' Purpose: Sets the caption picture
' Remarks:
'===============================================================================
Public Property Set CaptionPicture(ByVal New_Image As Picture)
    Set prvCaptionPicture = New_Image
    Set prvMenuForm.Picture = prvCaptionPicture
    PropertyChanged "CaptionPicture"
    prvMenuForm.UpdateIcon
End Property



Private Sub UserControl_Initialize()
    Set prvMenuItems = New MenuItems
    prvMenuItems.SetParentId 1000000000, 1000000000
    Set prvMenuForm = New frmMenu
    Set prvMenuForm.prvParent = Me
End Sub
Private Sub UserControl_InitProperties()
    Set prvIcon = frmMenu.Picture1.Picture
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set prvIcon = PropBag.ReadProperty("Icon", Nothing)
    Set prvCaptionPicture = PropBag.ReadProperty("CaptionPicture", Nothing)
    prvToolTip = PropBag.ReadProperty("ToolTip", "Tool Tip")
    prvCaption = PropBag.ReadProperty("Caption", "Caption")
    prvMenuForm.Caption = prvCaption
    Set prvMenuForm.Picture1.Picture = prvIcon
    prvMenuForm.UpdateIcon
End Sub
Private Sub UserControl_Resize()
    Width = 600
    Height = 600
End Sub

Private Sub UserControl_Terminate()
    DetachFromSysTray
    Unload FrmAbout
    Unload frmMenu
    Set prvMenuItems = Nothing
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", prvIcon, Nothing)
    Call PropBag.WriteProperty("CaptionPicture", prvCaptionPicture, Nothing)
    Call PropBag.WriteProperty("ToolTip", prvToolTip, "ToolTip")
    Call PropBag.WriteProperty("Caption", prvCaption, "Caption")
End Sub

'===============================================================================
' Name: Property Get ToolTip
' Input:
'
' Output:
'    String - Returns the SystemTray tooltip text
' Purpose: Gets the SystemTray tooltip text
' Remarks:
'===============================================================================
'MemberInfo=13,0,0,0
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Sets/gets the tooltip text of the system tray icon"
Attribute ToolTip.VB_ProcData.VB_Invoke_Property = ";Appearance"

    ToolTip = prvToolTip
    
End Property
'===============================================================================
' Name: Property Let ToolTip
' Input:
'   New_ToolTip as String - The SystemTray tooltip text
' Output:
'
' Purpose: Sets the SystemTray menu caption
' Remarks:
'===============================================================================
Public Property Let ToolTip(ByVal New_ToolTip As String)

    prvToolTip = New_ToolTip
    PropertyChanged "ToolTip"
    prvMenuForm.UpdateIcon
    
End Property

'===============================================================================
' Name: Property Get Caption
' Input:
'
' Output:
'    String - Returns the SystemTray menu caption
' Purpose: Gets SystemTray menu caption
' Remarks:
'===============================================================================
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/gets then vertical text shown by the menu"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_MemberFlags = "204"
    Caption = prvCaption
End Property
'===============================================================================
' Name: Property Set Caption
' Input:
'   ByVal New_Caption as string - The SystemTray menu caption
' Output:
'
' Purpose: Sets SystemTray menu caption
' Remarks:
'===============================================================================
Public Property Let Caption(ByVal New_Caption As String)
    prvCaption = New_Caption
    prvMenuForm.Caption = prvCaption
    PropertyChanged "Caption"
End Property
'===============================================================================
' Name: Procedure AtachToSysTray
' Input:
' Output:
' Purpose: Creates a system tray icon in the system tray
' Remarks:
'===============================================================================
'MemberInfo=0
Public Sub AtachToSysTray()
Attribute AtachToSysTray.VB_Description = "Creates a system tray icon in the system tray"
    
    prvMenuForm.CreateIcon
    
End Sub

'===============================================================================
' Name: Procedure Update
' Input:
' Output:
' Purpose: Updates the SysTray Icon
' Remarks:
'===============================================================================
'MemberInfo=0
Public Sub Update()
    
    prvMenuForm.UpdateIcon
    
End Sub

'===============================================================================
' Name: Procedure DetachFromSysTray
' Input:
' Output:
' Purpose: Removes the system tray icon from the system tray
' Remarks:
'===============================================================================
'MemberInfo=0
Public Sub DetachFromSysTray()
Attribute DetachFromSysTray.VB_Description = "Removes the system tray icon from the system tray"
    
    prvMenuForm.DeleteIcon
    Unload prvMenuForm
    
End Sub





Friend Function NotifyMenuItemSelect(ByVal Item As Long)

    RaiseEvent OnMenuItemSelect(Item)
    
End Function
Friend Function NotifyRightClick()
    RaiseEvent OnRightClick
End Function
Friend Function NotifyMouseMove()
    RaiseEvent OnMouseMove
End Function
Friend Function NotifyLeftClick()
    RaiseEvent OnLeftClick
End Function
Friend Function NotifyDBLClick()
    RaiseEvent OnDblClick
End Function
'===============================================================================
' Name: Procedure AboutBox
' Input:
' Output:
' Purpose: Shows the AboutBox
' Remarks:
'===============================================================================
Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Shows About Box"
Attribute AboutBox.VB_UserMemId = -552
    FrmAbout.Show vbModal
End Sub

'===============================================================================
' Name: Property Get MenuItems
' Input:
' Output:
'   MenuItems - The items of the SysTray Icon's menu
' Purpose: Gets the SysTray Icon's menu items
' Remarks:
'===============================================================================
Public Property Get MenuItems() As MenuItems

    Set MenuItems = prvMenuItems

End Property

