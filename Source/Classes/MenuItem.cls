VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
' Name: MenuItem
' Purpose: Menu item properties class.
' Functions:
' Properties:
' Methods:
' Author:Ciprian Sorlea
' Start: 15.01.2002
' Modified: 15.01.2002
'===============================================================================

'===============================================================================
' Name: Enum MenuItemType
'   mtMenuItem - Simple menu item or submenu
'   mtSeparator - Menu separator
' Purpose: Menu item type
' Remarks:
'===============================================================================
Public Enum MenuItemType
    mtMenuItem = 1
    mtSeparator = 5
End Enum


Private prvCaption As String
Private prvGrayed As Boolean
Private prvChecked As Boolean
Private prvID As Long
Private prvType As MenuItemType
Private prvIcon As StdPicture
Private prvSubMenus As MenuItems
Private prvIncrement As Long

'===============================================================================
' Name: Function IsSubMenu
' Input:
' Output:
'    Boolean - Returns true if current menu item is a submenu
' Purpose: Checks if if current menu item is a submenu
' Remarks:
'===============================================================================

Public Function IsSubMenu() As Boolean

    IsSubMenu = Not (prvSubMenus Is Nothing)
    
End Function

Private Sub Class_Terminate()

   Set prvSubMenus = Nothing
   Set prvIconSelected = Nothing
   Set prvIcon = Nothing

End Sub



'===============================================================================
' Name: Property Get Caption
' Input:
'
' Output:
'    String - Returns the menu item's caption
' Purpose: Gets menu item's caption
' Remarks:
'===============================================================================
Public Property Get Caption() As String
    Caption = prvCaption
End Property
'===============================================================================
' Name: Property Let Caption
' Input:
'    ByVal NewValue as String - The caption
' Output:
' Purpose: Sets menu item's caption
' Remarks:
'===============================================================================
Public Property Let Caption(ByVal NewValue As String)
    
    prvCaption = NewValue
    
End Property

'===============================================================================
' Name: Property Get Checked
' Input:
'
' Output:
'    Boolean - returns true if menu item is checked
' Purpose: Gets menu item's checked state
' Remarks:
'===============================================================================
Public Property Get Checked() As Boolean
    Checked = prvChecked
End Property

'===============================================================================
' Name: Property Let Checked
' Input:
'    ByVal NewValue as Boolean - Sets if menu item is checked or not
' Output:
' Purpose: Sets menu item's checked state
' Remarks:
'===============================================================================
Public Property Let Checked(ByVal NewValue As Boolean)

    prvChecked = NewValue
    
End Property


'===============================================================================
' Name: Property Get Grayed
' Input:
'
' Output:
'    Boolean - returns true if menuitem is grayed(disabled)
' Purpose: Gets menu item's enabled state
' Remarks:
'===============================================================================
Public Property Get Grayed() As Boolean

    Grayed = prvGrayed
    
End Property

'===============================================================================
' Name: Property Let Grayed
' Input:
'    BzVal NewValue as Boolean - menu itemþs  grayed(disabled) state
' Output:
' Purpose: Sets menu item's enabled state
' Remarks:
'===============================================================================
Public Property Let Grayed(ByVal NewValue As Boolean)

    prvGrayed = NewValue
    
End Property


'===============================================================================
' Name: Property Get ID
' Input:
'
' Output:
'    Long - menu item's ID
' Purpose: Gets menu item's id
' Remarks:
'===============================================================================
Public Property Get ID() As Long

    ID = prvID

End Property

Friend Property Let ID(ByVal lngID As Long)

    prvID = lngID

End Property

'===============================================================================
' Name: Property Get Icon
' Input:
' Output:
'    StdPicture - The icon of the menu item
' Purpose: Gets the menu item's icon
' Remarks:
'===============================================================================
Public Property Get Icon() As StdPicture

    Set Icon = prvIcon

End Property

'===============================================================================
' Name: Property Set Icon
' Input:
'    objIcon as StdPicture - The icon of the menu item'
' Output:
' Purpose: Sets the menu item's icon
' Remarks:
'===============================================================================
Public Property Set Icon(objIcon As StdPicture)

    Set prvIcon = objIcon

End Property

'===============================================================================
' Name: Property Get MenuType
' Input:
' Output:
'    MenuItemType - The type of the menu item
' Purpose: Gets the menu item's type
' Remarks:
'===============================================================================
Public Property Get MenuType() As MenuItemType

    MenuType = prvType

End Property

'===============================================================================
' Name: Property LGet MenuType
' Input:
'    ByVal enuMenuType as MenuItemType - The type of the menu item
' Output:
' Purpose: Sets the menu item's type
' Remarks:
'===============================================================================
Public Property Let MenuType(ByVal enuMenuType As MenuItemType)

    prvType = enuMenuType

End Property


'===============================================================================
' Name: Property Get MenuItems
' Input:
' Output:
'   MenuItems - The items of the current submenu
' Purpose: Gets the menu items of the current submenu
' Remarks:
'===============================================================================
Public Property Get MenuItems() As MenuItems

    If prvIncrement < 100 Then
        Err.Raise -8, "FlashTrayIcon", "Only 4 levels of menu items allowed"
        Exit Property
    End If
    
    If prvType = mtSeparator Then
        Err.Raise -6, "FlashTrayIcon", "Separators cannot have sub menus"
        Exit Property
    End If
    
    If prvSubMenus Is Nothing Then
        Set prvSubMenus = New MenuItems
        prvSubMenus.SetParentId prvID, prvIncrement
    End If
    
    Set MenuItems = prvSubMenus

End Property



Friend Property Get Increment() As Long

    Increment = prvIncrement

End Property

Friend Property Let Increment(ByVal lngIncrement As Long)

    prvIncrement = lngIncrement

End Property
