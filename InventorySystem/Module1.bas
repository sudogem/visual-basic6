Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public strPath As String


Public Sub Main()
strPath = App.Path & "\Database\InventoryDB.mdb"
'formLogin.Show
MainWindow.Show
End Sub

Public Sub CentreFormOnMdi(ByRef rfrmToBeCentred As Form, ByRef rfrmMDI As MDIForm)
' Purpose : centres form on MDI.
' Inputs  : rfrmToBeCentred - form to be centred.
'           rfrmMDI - MDI form to centre on.
' Outputs : none.
    With rfrmToBeCentred
        .Top = rfrmMDI.Top + (rfrmMDI.Height - .Height) / 2
        .Left = rfrmMDI.Left + (rfrmMDI.Width - .Width) / 2
    End With
End Sub

Public Sub CentreFormOnScreen(ByRef rfrmToBeCentred As Form)
' Purpose : centres form on screen.
' Inputs  : rfrmToBeCentred - form to be centred.
' Outputs : none.
    With rfrmToBeCentred
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Public Sub ShowError(message As String)
' Purpose : create my own error message box by just passing message.
    formError.errorMessage (message)
    formError.Show vbModal
End Sub

Public Sub HideContainers()
    MainWindow.containerBottom.Visible = False
    MainWindow.contanerMiddle.Visible = False
    MainWindow.containerTop.Visible = False
End Sub

Public Sub ShowContainers()
    MainWindow.containerTop.Visible = True
    MainWindow.contanerMiddle.Visible = True
    MainWindow.containerBottom.Visible = True
End Sub

'FOR SUB MENUS
Public Sub AddUserClicked()
    MainWindow.EditUser.Picture = MainWindow.submenu_ImageList2.ListImages(1).Picture
    MainWindow.DeleteUser.Picture = MainWindow.submenu_ImageList3.ListImages(1).Picture
End Sub

Public Sub EditUserClicked()
    MainWindow.AddUser.Picture = MainWindow.submenu_ImageList1.ListImages(1).Picture
    MainWindow.DeleteUser.Picture = MainWindow.submenu_ImageList3.ListImages(1).Picture
End Sub

Public Sub DeleteUserClicked()
    MainWindow.AddUser.Picture = MainWindow.submenu_ImageList1.ListImages(1).Picture
    MainWindow.EditUser.Picture = MainWindow.submenu_ImageList2.ListImages(1).Picture
End Sub


Public Sub AddItemsAndPackagesClicked()
    MainWindow.Supply.Picture = MainWindow.submenu_ImageList5.ListImages(1).Picture
    MainWindow.InventoryUpdate.Picture = MainWindow.submenu_ImageList6.ListImages(1).Picture
    MainWindow.EditItemsAndPackages.Picture = MainWindow.submenu_ImageList9.ListImages(1).Picture
    MainWindow.DeleteItemsAndPackages.Picture = MainWindow.submenu_ImageList10.ListImages(1).Picture
End Sub

Public Sub EditItemsAndPackagesClicked()
    MainWindow.Supply.Picture = MainWindow.submenu_ImageList5.ListImages(1).Picture
    MainWindow.InventoryUpdate.Picture = MainWindow.submenu_ImageList6.ListImages(1).Picture
    MainWindow.AddItemsAndPackages.Picture = MainWindow.submenu_ImageList4.ListImages(1).Picture
    MainWindow.DeleteItemsAndPackages.Picture = MainWindow.submenu_ImageList10.ListImages(1).Picture
End Sub

Public Sub DeleteItemsAndPackagesClicked()
    MainWindow.Supply.Picture = MainWindow.submenu_ImageList5.ListImages(1).Picture
    MainWindow.InventoryUpdate.Picture = MainWindow.submenu_ImageList6.ListImages(1).Picture
    MainWindow.AddItemsAndPackages.Picture = MainWindow.submenu_ImageList4.ListImages(1).Picture
    MainWindow.EditItemsAndPackages.Picture = MainWindow.submenu_ImageList9.ListImages(1).Picture
End Sub

Public Sub SupplyClicked()
    MainWindow.InventoryUpdate.Picture = MainWindow.submenu_ImageList6.ListImages(1).Picture
    MainWindow.AddItemsAndPackages.Picture = MainWindow.submenu_ImageList4.ListImages(1).Picture
    MainWindow.EditItemsAndPackages.Picture = MainWindow.submenu_ImageList9.ListImages(1).Picture
    MainWindow.DeleteItemsAndPackages.Picture = MainWindow.submenu_ImageList10.ListImages(1).Picture
End Sub

Public Sub InventoryUpdateClicked()
    MainWindow.Supply.Picture = MainWindow.submenu_ImageList5.ListImages(1).Picture
    MainWindow.AddItemsAndPackages.Picture = MainWindow.submenu_ImageList4.ListImages(1).Picture
    MainWindow.EditItemsAndPackages.Picture = MainWindow.submenu_ImageList9.ListImages(1).Picture
    MainWindow.DeleteItemsAndPackages.Picture = MainWindow.submenu_ImageList10.ListImages(1).Picture
End Sub


Public Sub DailySalesClicked()
    MainWindow.SupplyReports.Picture = MainWindow.submenu_ImageList8.ListImages(1).Picture
End Sub

Public Sub SupplyReportsClicked()
    MainWindow.DailySales.Picture = MainWindow.submenu_ImageList7.ListImages(1).Picture
End Sub


'FOR MAIN MENUS
Public Sub UsersMenuClicked()
    MainWindow.InventoryMenu.Picture = MainWindow.ImageList2.ListImages(1).Picture
    MainWindow.ReportsMenu.Picture = MainWindow.ImageList3.ListImages(1).Picture
    MainWindow.BuyMenu.Picture = MainWindow.ImageList4.ListImages(1).Picture
End Sub

Public Sub InventoryMenuClicked()
    MainWindow.UsersMenu.Picture = MainWindow.ImageList1.ListImages(1).Picture
    MainWindow.ReportsMenu.Picture = MainWindow.ImageList3.ListImages(1).Picture
    MainWindow.BuyMenu.Picture = MainWindow.ImageList4.ListImages(1).Picture
End Sub

Public Sub ReportsMenuClicked()
    MainWindow.UsersMenu.Picture = MainWindow.ImageList1.ListImages(1).Picture
    MainWindow.InventoryMenu.Picture = MainWindow.ImageList2.ListImages(1).Picture
    MainWindow.BuyMenu.Picture = MainWindow.ImageList4.ListImages(1).Picture
End Sub

Public Sub BuyMenuClicked()
    MainWindow.UsersMenu.Picture = MainWindow.ImageList1.ListImages(1).Picture
    MainWindow.InventoryMenu.Picture = MainWindow.ImageList2.ListImages(1).Picture
    MainWindow.ReportsMenu.Picture = MainWindow.ImageList3.ListImages(1).Picture
End Sub


