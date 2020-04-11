VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectTemplateDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_Opened(CancelForwarding As Boolean)

    'Declare the following variables and set aside storage space
    Dim objDocMenu As HMIMenu
    Dim objMenuItem As HMIMenuItem
    
    'if there is an error, we are just going to ignore it and continue
    On Error Resume Next
    
    'Configure the Menu Toolbar at position 1 and Menu Item at position 1
    Set objDocMenu = Application.CustomMenus.InsertMenu(1, "VBA_Tasks", "VBA Macros")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(1, "Trend_Creation", "Create Trend")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(2, "Object_Alignment", "Align Objects")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(3, "IOField_Sizing", "Size I/O Fields")
    'Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(4, "Testing", "Test Script")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(4, "Add_Navigation", "Add Navigation")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(5, "Pipe_Tool", "Pipe Tool")
    Set objMenuItem = objDocMenu.MenuItems.InsertMenuItem(6, "Mass_Object_Tool", "Mass Object Tool")
    
    'assign the vba macro/subroutine to run
    With Application.CustomMenus("VBA_Tasks")
        .MenuItems("Trend_Creation").Macro = "Trend_Creation"
        .MenuItems("Object_Alignment").Macro = "Object_Alignment"
        .MenuItems("IOField_Sizing").Macro = "IOField_Sizing"
        '.MenuItems("Testing").Macro = "Testing"
        .MenuItems("Add_Navigation").Macro = "Add_Navigation"
        .MenuItems("Pipe_Tool").Macro = "Pipe_Tool"
        .MenuItems("Mass_Object_Tool").Macro = "Mass_Object_Tool"
    End With

End Sub
