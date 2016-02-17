Attribute VB_Name = "Module2"
Public Sub SetAllSubmenuToDefault()
     
    MainWindow.AddUser.Picture = MainWindow.submenu_ImageList1.ListImages(1).Picture
    MainWindow.EditUser.Picture = MainWindow.submenu_ImageList2.ListImages(1).Picture
    MainWindow.DeleteUser.Picture = MainWindow.submenu_ImageList3.ListImages(1).Picture
    MainWindow.AddItemsAndPackages = MainWindow.submenu_ImageList4.ListImages(1).Picture
    MainWindow.EditItemsAndPackages = MainWindow.submenu_ImageList9.ListImages(1).Picture
    MainWindow.DeleteItemsAndPackages = MainWindow.submenu_ImageList10.ListImages(1).Picture
    MainWindow.Supply.Picture = MainWindow.submenu_ImageList5.ListImages(1).Picture
    MainWindow.InventoryUpdate.Picture = MainWindow.submenu_ImageList6.ListImages(1).Picture
    MainWindow.DailySales.Picture = MainWindow.submenu_ImageList7.ListImages(1).Picture
    MainWindow.SupplyReports.Picture = MainWindow.submenu_ImageList8.ListImages(1).Picture
End Sub


Public Sub HideAllSubmenuContainer()
    MainWindow.DefaultSubMenuContainer.Visible = True
    MainWindow.UsersContainer.Visible = False
    MainWindow.InventorySubmenuContainer.Visible = False
    MainWindow.ReportsContainer.Visible = False
End Sub

Public Sub HideSubmenuForms()
    'formAddUser.Hide
    'formEditUser.Hide
    'formDeleteUser.Hide
    'formAddItemsAndPackages.Hide
    'formEditItemsAndPackages.Hide
    'formDeleteItemsAndPackages.Hide
    'formSupply.Hide
    'formInventoryUpdate.Hide
    'formBuy.Hide
    
    Unload formAddUser
    Unload formEditUser
    Unload formDeleteUser
    Unload formAddItemsAndPackages
    Unload formEditItemsAndPackages
    Unload formDeleteItemsAndPackages
    Unload formSupply
    Unload formInventoryUpdate
    Unload formBuy
    Unload formSupplierReport
    Unload formSalesReport
End Sub
