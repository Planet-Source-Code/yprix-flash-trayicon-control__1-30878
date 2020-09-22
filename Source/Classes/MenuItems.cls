VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
' Name: MenuItems
' Purpose: Menu item collection class
' Functions:
' Properties:
' Methods:
' Author:Ciprian Sorlea
' Start: 15.01.2002
' Modified: 15.01.2002
'===============================================================================

'===============================================================================
' Name: Enum Constants
'   MaxSubMenus - Maximum number of menuitems in a menu or submenu
' Purpose: Diferent constants used in project
' Remarks:
'===============================================================================
Public Enum Constants
    MaxSubMenus = 25
End Enum

Private prvParentId As Long
Private prvIncrement As Long
Private prvMenuCnt As Integer
Private prvMenuItems As Collection

'===============================================================================
' Name: Function Add
' Input:
'   ByVal mnuType As MenuItemType - the type of the menu item
'   ByVal mnuCaption As String - the caption of the menu item
'   ByVal mnuPosition As Integer - the position of the menu item<br> -1 : at the end<br>-2 : at the beginning<br>otherways : before the mnuPosition menu item
' Output:
'   MenuItem - the new menu item created
' Purpose: Adds a new menu item to the system tray icon's menu
' Remarks:
'===============================================================================
Public Function Add(ByVal mnuType As MenuItemType, _
                    Optional ByVal mnuCaption As String = " ", _
                    Optional ByVal mnuPosition As Integer = -1) As MenuItem
    


    If Count = MaxSubMenus Then
        Err.Raise -7, "FlashTrayIcon", "Only " + CStr(MaxSubMenus) + " menu items allowed"
        Exit Function
    End If
    Dim newItem As MenuItem
    
    Select Case mnuType
    
        Case mtMenuItem
            
            If Trim(mnuCaption) = "" Then
                Err.Raise -1, "FlashTrayIcon", "Caption cannot be empty"
                Exit Function
            End If
            
            Set newItem = New MenuItem
            newItem.Caption = mnuCaption
            newItem.MenuType = mtMenuItem
            
            
       
            
        Case mtSeparator
        
            Set newItem = New MenuItem
            newItem.MenuType = mtSeparator
        
        Case Else
            Err.Raise -2, "FlashTrayIcon", "No such MenuType supported"
            Exit Function
            
    End Select
    
On Error GoTo ErrorHandler

    prvMenuCnt = prvMenuCnt + 1
    newItem.ID = prvParentId + prvIncrement * prvMenuCnt
    newItem.Increment = prvIncrement
    Debug.Print newItem.ID
    Select Case mnuPosition
        
        Case -1
            prvMenuItems.Add newItem, CStr(newItem.ID)
            
        Case 0
            prvMenuItems.Add newItem, CStr(newItem.ID), 1
            
        Case Else
            If mnuPosition <= prvMenuItems.Count Then
                prvMenuItems.Add newItme, CStr(newItem.ID), mnuPosition
            Else
                prvMenuItems.Add newItem
            End If
        
    End Select
    
    Set Add = newItem
    
    Exit Function
    
ErrorHandler:
    Err.Raise -999, "FlashTrayIcon", "Unspecified error:" + vbNewLine + Err.Description
    
End Function



'===============================================================================
' Name: Property Get Count
' Input:
' Output:
'   Integer -  the number of menu items
' Purpose: Returns the number of items in the current menu or submenu
' Remarks:
'===============================================================================
Public Property Get Count() As Integer
    
    Count = prvMenuItems.Count
    
End Property

Friend Function GetMenu() As Long
    
    Dim hMenu As Long
    hMenu = CreatePopupMenu()
    For I = 1 To Count
        With Item(I)
            If .MenuType = mtSeparator Then
                 AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
            Else
                style = 0&
                If .Checked = True Then
                    style = style + MF_CHECKED
                End If
                
                If .Grayed = True Then
                    style = style + MF_GRAYED
                End If
                

                If .IsSubMenu Then
                    AppendMenu hMenu, style + MF_POPUP, .MenuItems.GetMenu, .Caption
                Else
                    AppendMenu hMenu, style, .ID, .Caption
                End If
                
                'SetMenuItemBitmaps hMenu, I, MF_BYPOSITION, .Icon, .IconSelected
            End If
            
        End With
    Next
    
    GetMenu = hMenu
End Function

'===============================================================================
' Name: Property Get Item
' Input:
'   ByVal Index as Integer -  the index of the menu item
' Output:
'   MenuItem - the selected menu item
' Purpose: Returns the menu item on the [index] position in the current menu or submenu
' Remarks:
'===============================================================================
Public Property Get Item(ByVal Index As Integer) As MenuItem
Attribute Item.VB_UserMemId = 0
    
    If Index > prvMenuItems.Count Or Index < 1 Then
        Err.Raise -5, "FlashTrayIcon", "Bad MenuItem index!!!"
        Exit Property
    End If
    Set Item = prvMenuItems(Index)
    
End Property

'===============================================================================
' Name: Procedure Add
' Input:
'   ByVal Index As Integer - the index of the menu item
' Output:
' Purpose: Removes the selected item from the menu
' Remarks:
'===============================================================================
Public Sub Remove(ByVal Index As Integer)
    
    If Index > prvMenuItems.Count Or Index < 1 Then
        Err.Raise -5, "FlashTrayIcon", "Bad MenuItem index!!!"
        Exit Sub
    End If
    prvMenuItems.Remove (Index)
    
End Sub


Friend Sub SetParentId(ByVal Value As Long, ByVal ParentIncrement As Long)

    prvParentId = Value
    prvIncrement = ParentIncrement / 100
    
End Sub

Private Sub Class_Initialize()

    Set prvMenuItems = New Collection
    
End Sub


Private Sub Class_Terminate()

    Set prvMenuItems = Nothing
    
End Sub



