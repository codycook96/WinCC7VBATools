VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mass_Object_Tool_Form 
   Caption         =   "Mass Object Tool"
   ClientHeight    =   7575
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5955
   OleObjectBlob   =   "Mass_Object_Tool_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mass_Object_Tool_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pages() As Document
Dim types() As String
Dim objects() As HMIObject
Dim properties() As HMIProperties
Dim pagesSize As Integer
Dim typesSize As Integer
Dim objectsSize As Integer
Dim propertiesSize As Integer

Dim selectEvents As Boolean
Dim valueExpanded As Boolean

Private Sub UserForm_Initialize()
    Lock_After_Page 0
End Sub

'============== General Page ================'
'------------ General Functions -------------'
Sub Set_ListBox(ByRef ctrl As Control, ByRef newList() As String)
    Dim i, j, k
    Dim newListSize, newListCols, temp
    Dim maxCharCount, charWidth, columnWidths
    
    charWidth = 7
    
    On Error Resume Next
    
    temp = UBound(newList, 2)
    
    ctrl.Clear
    
    If Err.Number = 0 Then
    '2D Array
        newListSize = UBound(newList, 2) - LBound(newList, 2) + 1
        newListCols = UBound(newList, 1) - LBound(newList, 1) + 1
        ctrl.ColumnCount = newListCols
        For j = 0 To newListCols - 1
            maxCharCount = 0
            For i = 0 To newListSize - 1
                If j = 0 Then
                    ctrl.AddItem
                End If
                If Len(newList(j, i)) > maxCharCount Then
                    maxCharCount = Len(newList(j, i))
                End If
                ctrl.list(i, j) = newList(j, i)
            Next i
                
                columnWidths = columnWidths & (maxCharCount * charWidth) + 10 & " pt" & "; "
                
        Next j
        ctrl.columnWidths = columnWidths
    Else
    '1D Array
        newListSize = UBound(newList) - LBound(newList) + 1
        newListCols = 0
        ctrl.ColumnCount = 1
        For i = 0 To newListSize - 1
            ctrl.AddItem
            ctrl.list(i) = newList(i)
        Next
    End If
    
    Err.Clear
End Sub

Sub Sort_ListBox(ByRef ctrl As Control)
    Dim i, j, k, m, n
    Dim newListSize, newListCols
    Dim temp()
    Dim errCheck
    Dim previousEqual
    
    On Error Resume Next
    
    If ctrl.ColumnCount > 1 Then
    '2D Array
        newListSize = ctrl.ListCount
        newListCols = ctrl.ColumnCount
        ReDim temp(newListCols)
        For k = 0 To newListCols - 1
            For i = 0 To newListSize - 2
                For j = i + 1 To newListSize - 1
                    previousEqual = True
                    For n = 0 To k - 1
                        If ctrl.list(i, n) <> ctrl.list(j, n) Then
                            previousEqual = False
                        End If
                    Next n
                    If (ctrl.list(i, k) > ctrl.list(j, k)) And previousEqual Then
                        For m = 0 To newListCols - 1
                            temp(m) = ctrl.list(j, m)
                            ctrl.list(j, m) = ctrl.list(i, m)
                            ctrl.list(i, m) = temp(m)
                        Next m
                    End If
                Next j
            Next i
        Next k
    Else
    '1D Array
        ReDim temp(0)
        newListSize = ctrl.ListCount
        For i = 0 To newListSize - 2
            For j = i + 1 To newListSize - 1
                If ctrl.list(i) > ctrl.list(j) Then
                    temp(0) = ctrl.list(j)
                    ctrl.list(j) = ctrl.list(i)
                    ctrl.list(i) = temp(0)
                End If
            Next j
        Next i
    End If
    
    Err.Clear
End Sub
'-------- END General Functions END ---------'
'========== END General Page END ============'

'================ MultiPage ================='


'================ MultiPage ================='
'----------- MultiPage Interface ------------'

'------- END MultiPage Interface END---------'

'----------- MultiPage Functions ------------'
Sub Lock_After_Page(ByVal pageNum As Integer)
    Dim i, j
    Dim optionsList(3) As String
    optionsList(0) = "Static"
    optionsList(1) = "Dynamic"
    optionsList(2) = "Update Cycle"
    optionsList(3) = "Indirect"
    
    For i = 0 To MultiPage.pages.Count - 1
        If i <= pageNum Then
            MultiPage.pages(i).Enabled = True
        Else
            MultiPage.pages(i).Enabled = False
            
            For j = 0 To MultiPage.pages(i).Controls.Count - 1
                If TypeName(MultiPage.pages(i).Controls.Item(j)) = "ListBox" Then
                    
                    MultiPage.pages(i).Controls.Item(j).Clear
                End If
            Next j
            
        End If
    Next i
    
    Set_ListBox ListBox_Options, optionsList
    ListBox_Options.Selected(0) = True
    
End Sub
'------- END MultiPage Functions END---------'
'============ END MultiPage END ============='

'================ Pages Page ================'
'------------- Pages Interface --------------'
Private Sub Button_Pages_CurrentPage_Click()
    Dim pagesList() As String
    
    ReDim pagesList(0)
    
    pagesList(0) = ActiveDocument.Name
    
    Set_ListBox ListBox_Pages, pagesList
    
    Lock_After_Page 0
End Sub

Private Sub Button_Pages_AllPages_Click()
    Dim pagesList() As String
    
    All_Pages pagesList
    
    Set_ListBox ListBox_Pages, pagesList
    Sort_ListBox ListBox_Pages
    
    Lock_After_Page 0
End Sub

Private Sub Button_Pages_OpenedPages_Click()
    Dim i
    Dim pagesList() As String
    Dim pagesListSize As Integer
    
    pagesListSize = 0
    
    For i = 3 To Application.Documents.Count
        ReDim Preserve pagesList(pagesListSize)
        pagesList(pagesListSize) = Application.Documents.Item(i).Name
        pagesListSize = pagesListSize + 1
    Next i
    
    Set_ListBox ListBox_Pages, pagesList
    Sort_ListBox ListBox_Pages
    
    Lock_After_Page 0
End Sub

Private Sub Button_Pages_RemoveSelected_Click()
    Dim i, listboxSize
    listboxSize = ListBox_Pages.ListCount - 2
    
    For i = 0 To ListBox_Pages.ListCount - 1
        If i < ListBox_Pages.ListCount Then
            If ListBox_Pages.Selected(i) Then
                ListBox_Pages.RemoveItem (i)
                listboxSize = listboxSize - 1
                i = i - 1
            End If
        End If
    Next i
    
    Lock_After_Page 0
End Sub

Private Sub Button_Pages_Clear_Click()
    ListBox_Pages.Clear
    
    Lock_After_Page 0
End Sub

Private Sub Button_Pages_SetPages_Click()
    Set_Pages
End Sub

Private Sub ListBox_Pages_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, ListBox_Pages
End Sub

'--------- END Pages Interface END ----------'

'------------- Pages Functions --------------'
Sub Set_ListBox_Pages(ByRef newPageList() As String)
    Dim i, newPageListSize

    newPageListSize = UBound(newPageList) - LBound(newPageList) + 1
    
    'ListBox_Pages.
    For i = 0 To newPageListSize - 1
        ListBox_Pages.AddItem newPageList(i)
    Next
End Sub

Sub Set_Pages()
    Dim i, j
    Dim continue
    Dim unopenedCount
    
    unopenedCount = ListBox_Pages.ListCount
    
    ReDim unopened(unopenedCount)
    
    For i = 0 To ListBox_Pages.ListCount - 1
        unopened(i) = True
        For j = 1 To Application.Documents.Count
            If Application.Documents.Item(j).Name = ListBox_Pages.list(i) Then
                unopened(i) = False
                unopenedCount = unopenedCount - 1
            End If
        Next j
    Next i
    
    If unopenedCount > 10 Then
        continue = MsgBox("You are about to open " & unopenedCount & " pages. This could take a minute.", vbOKCancel, "Warning")
    Else
        continue = vbOK
    End If
    
    If continue = vbOK Then
        pagesSize = 0
        ReDim pages(pagesSize)
    
        For i = 0 To ListBox_Pages.ListCount - 1
            If unopened(i) Then
                Application.Documents.Open (ListBox_Pages.list(i))
            End If
            
            ReDim Preserve pages(pagesSize)
            
            For j = 1 To Application.Documents.Count
                If ListBox_Pages.list(i) = Application.Documents.Item(j).Name Then
                    Set pages(pagesSize) = Application.Documents.Item(j)
                End If
            Next j

            pagesSize = pagesSize + 1
        Next i
    End If
    
    If pagesSize > 0 Then
        Lock_After_Page 1
    End If
    
End Sub


Sub All_Pages(ByRef pageList() As String)
    'Get Path to Screen Folder
    Dim path, colPath, sizeColPath, i
    colPath = Split(ActiveDocument.path, "\")
    sizeColPath = UBound(colPath) - LBound(colPath)
    For i = 0 To sizeColPath - 1
        path = path + colPath(i) + "\"
    Next i
    
    'Add Screen Names to ListBox
    Dim objFSO, objFolder, colFiles, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(path)
    Set colFiles = objFolder.Files
    
    ReDim pageList(0)
    Dim pageListSize
    pageListSize = 0
    
    For Each objFile In colFiles
        If UCase(Right(objFile.Name, 3)) = "PDL" Then
            'Add further filters below like so...
            'If Left(objFile.Name, 2) <> "pw" And Left(objFile.Name, 2) <> "tr" And _
            'Left(objFile.Name, 2) <> "xx" And Left(objFile.Name, 2) <> "zz" And _
            'Left(objFile.Name, 2) <> "fp" Then
                ReDim Preserve pageList(pageListSize)
                pageList(pageListSize) = objFile.Name
                pageListSize = pageListSize + 1
            'End If
        End If
    Next objFile
End Sub
'--------- END Pages Functions END ----------'
'============ END Pages Page END ============'


'================ Types Page ================'
'------------- Types Interface --------------'
Private Sub Button_Types_CurrentSelection_Click()
    Dim typeList() As String
    
    Selected_Types typeList
    
    Set_ListBox ListBox_Types, typeList
    Sort_ListBox ListBox_Types
    
    Lock_After_Page 1
End Sub

Private Sub Button_Types_AllTypes_Click()
    Dim typeList() As String

    All_Types typeList
    
    Set_ListBox ListBox_Types, typeList
    Sort_ListBox ListBox_Types
    
    Lock_After_Page 1
End Sub

Private Sub Button_Types_CurrentPage_Click()
    Dim typeList() As String
    
    CurrentPage_Types typeList
    
    Set_ListBox ListBox_Types, typeList
    Sort_ListBox ListBox_Types
    
    Lock_After_Page 1
End Sub

Private Sub Button_Types_RemoveSelected_Click()
    Dim i, listboxSize
    listboxSize = ListBox_Types.ListCount - 2
    
    For i = 0 To ListBox_Types.ListCount - 1
        If i < ListBox_Types.ListCount Then
            If ListBox_Types.Selected(i) Then
                ListBox_Types.RemoveItem (i)
                listboxSize = listboxSize - 1
                i = i - 1
            End If
        End If
    Next i
    
    Lock_After_Page 1
End Sub

Private Sub Buttons_Types_Clear_Click()
    ListBox_Types.Clear
    
    Lock_After_Page 1
End Sub

Private Sub Button_Types_SetTypes_Click()
    Set_Types
End Sub

Private Sub ListBox_Types_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, ListBox_Types
End Sub
'--------- END Types Interface END ----------'

'------------- Types Functions --------------'
Sub Set_TypesList(ByRef typeList() As String)
    Dim typeListSize
    
    ListBox_Types.Clear
    
    ReDim types(0)
    typesSize = 0

    MultiPage.pages(2).Enabled = False

    ListBox_Properties.Clear
    
    typeListSize = (UBound(typeList, 2) - LBound(typeList, 2)) + 1

    For i = 0 To typeListSize - 1
        ListBox_Types.AddItem
        ListBox_Types.list(i, 0) = typeList(0, i)
        ListBox_Types.list(i, 1) = typeList(1, i)
    Next i
    
End Sub

Sub Set_Types()
    Dim i
    
    ReDim types(1, 0)

    For i = 0 To ListBox_Types.ListCount - 1
        ReDim Preserve types(1, typesSize)
        types(0, i) = ListBox_Types.list(i, 0)
        types(1, i) = ListBox_Types.list(i, 1)
        typesSize = typesSize + 1
    Next i
    
    If typesSize > 0 Then
        Lock_After_Page 2
    End If
End Sub

Sub All_Types(ByRef typeList() As String)
    Dim i, j, k
    Dim pageSelected
    Dim typeIncluded
    Dim typeListSize As Integer
    typeListSize = 0
    ReDim typeList(1, typeListSize)
    
    For i = 0 To pagesSize - 1
        For j = 1 To pages(i).HMIObjects.Count
            typeIncluded = False
            For k = 0 To typeListSize - 1
                If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" And _
                    typeList(0, k) = "FP" Then
                    If pages(i).HMIObjects.Item(j).properties.Item(3) = typeList(1, k) Then
                        typeIncluded = True
                    End If
                ElseIf pages(i).HMIObjects.Item(j).Type = typeList(1, k) Then
                    typeIncluded = True
                End If
            Next k
            If Not (typeIncluded) Then
                ReDim Preserve typeList(1, typeListSize)
                If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" Then
                    typeList(0, typeListSize) = "FP"
                    typeList(1, typeListSize) = pages(i).HMIObjects.Item(j).properties.Item(3)
                Else
                    typeList(0, typeListSize) = "HMI"
                    typeList(1, typeListSize) = pages(i).HMIObjects.Item(j).Type
                End If
                typeListSize = typeListSize + 1
            End If
        Next j
    Next i
End Sub

Sub Selected_Types(ByRef typeList() As String)
    Dim i, j, k
    Dim pageSelected
    Dim typeIncluded
    Dim typeListSize As Integer
    typeListSize = 0
    ReDim typeList(1, typeListSize)
    
    For i = 0 To pagesSize - 1
        For j = 1 To pages(i).HMIObjects.Count
            typeIncluded = False
            For k = 0 To typeListSize - 1
                If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" And _
                    typeList(0, m) = "FP" Then
                    If pages(i).HMIObjects.Item(j).properties.Item(3) = typeList(1, k) Then
                        typeIncluded = True
                    End If
                ElseIf pages(i).HMIObjects.Item(j).Type = typeList(1, k) Then
                    typeIncluded = True
                End If
            Next k
            If Not (typeIncluded) Then
                If pages(i).HMIObjects.Item(j).Selected Then
                    ReDim Preserve typeList(1, typeListSize)
                    If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" Then
                        typeList(0, typeListSize) = "FP"
                        typeList(1, typeListSize) = pages(i).HMIObjects.Item(j).properties.Item(3)
                    Else
                        typeList(0, typeListSize) = "HMI"
                        typeList(1, typeListSize) = pages(i).HMIObjects.Item(j).Type
                    End If
                    typeListSize = typeListSize + 1
                End If
            End If
        Next j
    Next i
End Sub

Sub CurrentPage_Types(ByRef typeList() As String)
Dim k, m
    Dim pageSelected
    Dim typeIncluded
    Dim typeListSize As Integer
    typeListSize = 0
    ReDim typeList(1, typeListSize)
    
    For k = 1 To Application.ActiveDocument.HMIObjects.Count
        typeIncluded = False
        For m = 0 To typeListSize - 1
            If Application.ActiveDocument.HMIObjects.Item(k).Type = "HMIFaceplateObject" And _
                typeList(0, m) = "FP" Then
                If Application.ActiveDocument.HMIObjects.Item(k).properties.Item(3) = typeList(1, m) Then
                    typeIncluded = True
                End If
            ElseIf Application.ActiveDocument.HMIObjects.Item(k).Type = typeList(1, m) Then
                typeIncluded = True
            End If
        Next m
        If Not (typeIncluded) Then
            ReDim Preserve typeList(1, typeListSize)
            If Application.ActiveDocument.HMIObjects.Item(k).Type = "HMIFaceplateObject" Then
                typeList(0, typeListSize) = "FP"
                typeList(1, typeListSize) = Application.ActiveDocument.HMIObjects.Item(k).properties.Item(3)
            Else
                typeList(0, typeListSize) = "HMI"
                typeList(1, typeListSize) = Application.ActiveDocument.HMIObjects.Item(k).Type
            End If
            typeListSize = typeListSize + 1
        End If
    Next k
End Sub

'--------- END Types Interface END ----------'
'============ END Types Page END ============'

'=============== Objects Page ==============='
'------------- Objects Interface ------------'
Private Sub Button_Objects_CurrentSelection_Click()
    Dim objectList() As String
    
    Selected_Objects objectList
    
    Set_ListBox ListBox_Objects, objectList
    Sort_ListBox ListBox_Objects
    
    Lock_After_Page 2
End Sub

Private Sub Button_Objects_AllObjects_Click()
    Dim objectList() As String
    
    All_Objects objectList
    
    Set_ListBox ListBox_Objects, objectList
    Sort_ListBox ListBox_Objects
    
    Lock_After_Page 2
End Sub

Private Sub Button_Objects_CurrentPage_Click()
    Dim objectList() As String
    
    CurrentPage_Objects objectList
    
    Set_ListBox ListBox_Objects, objectList
    Sort_ListBox ListBox_Objects
    
    Lock_After_Page 2
End Sub

Private Sub Button_Objects_Clear_Click()
    ListBox_Objects.Clear
    
    Lock_After_Page 2
End Sub

Private Sub Button_Objects_RemoveSelected_Click()
    Dim i, listboxSize
    listboxSize = ListBox_Objects.ListCount - 2
    
    For i = 0 To ListBox_Objects.ListCount - 1
        If i < ListBox_Objects.ListCount Then
            If ListBox_Objects.Selected(i) Then
                ListBox_Objects.RemoveItem (i)
                listboxSize = listboxSize - 1
                i = i - 1
            End If
        End If
    Next i
    
    Lock_After_Page 2
End Sub

Private Sub Button_Objects_SetObjects_Click()
    Set_Objects
End Sub

Private Sub ListBox_Objects_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, ListBox_Objects
End Sub
'--------- END Objects Interface END --------'

'------------- Objects Functions ------------'
Sub All_Objects(ByRef objectList() As String)
    Dim i, j, k
    Dim objectListSize
    
    objectListSize = 0
    
    ReDim objectList(1, objectListSize)
    
    For i = 0 To pagesSize - 1
        For j = 1 To pages(i).HMIObjects.Count
            For k = 0 To typesSize - 1
                If types(0, k) = "FP" Then
                    If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" And pages(i).HMIObjects.Item(j).properties.Item(3) = types(1, k) Then
                        ReDim Preserve objectList(1, objectListSize)
                        objectList(0, objectListSize) = pages(i).HMIObjects.Item(j).ObjectName
                        objectList(1, objectListSize) = pages(i).Name
                        objectListSize = objectListSize + 1
                    End If
                ElseIf types(0, k) = "HMI" Then
                    If pages(i).HMIObjects.Item(j).Type = types(1, k) Then
                        ReDim Preserve objectList(1, objectListSize)
                        objectList(0, objectListSize) = pages(i).HMIObjects.Item(j).ObjectName
                        objectList(1, objectListSize) = pages(i).Name
                        objectListSize = objectListSize + 1
                    End If
                End If
            Next k
        Next j
    Next i
End Sub

Sub Selected_Objects(ByRef objectList() As String)
    Dim i, j, k
    Dim objectListSize
    
    objectListSize = 0
    
    ReDim objectList(1, objectListSize)
    
    For i = 0 To pagesSize - 1
        For j = 1 To pages(i).HMIObjects.Count
            For k = 0 To typesSize - 1
                If types(0, k) = "FP" Then
                        If pages(i).HMIObjects.Item(j).Type = "HMIFaceplateObject" And _
                        pages(i).HMIObjects.Item(j).properties.Item(3) = types(1, k) And _
                        (pages(i).HMIObjects.Item(j).Selected) Then
                        
                        ReDim Preserve objectList(1, objectListSize)
                        objectList(0, objectListSize) = pages(i).HMIObjects.Item(j).ObjectName
                        objectList(1, objectListSize) = pages(i).Name
                        objectListSize = objectListSize + 1
                        
                    End If
                ElseIf types(0, k) = "HMI" Then
                    If (pages(i).HMIObjects.Item(j).Type = types(1, k)) And (pages(i).HMIObjects.Item(j).Selected) Then
                        
                        ReDim Preserve objectList(1, objectListSize)
                        objectList(0, objectListSize) = pages(i).HMIObjects.Item(j).ObjectName
                        objectList(1, objectListSize) = pages(i).Name
                        objectListSize = objectListSize + 1
                        
                    End If
                End If
            Next k
        Next j
    Next i
End Sub

Sub CurrentPage_Objects(ByRef objectList() As String)
    Dim i, j, k
    Dim objectListSize
    
    objectListSize = 0
    
    ReDim objectList(1, objectListSize)
    
    For j = 1 To Application.ActiveDocument.HMIObjects.Count
        For k = 0 To typesSize - 1
            If types(0, k) = "FP" Then
                    If Application.ActiveDocument.HMIObjects.Item(j).Type = "HMIFaceplateObject" And Application.ActiveDocument.HMIObjects.Item(j).properties.Item(3) = types(1, k) Then
                    ReDim Preserve objectList(1, objectListSize)
                    objectList(0, objectListSize) = Application.ActiveDocument.HMIObjects.Item(j).ObjectName
                    objectList(1, objectListSize) = Application.ActiveDocument.Name
                    objectListSize = objectListSize + 1
                End If
            ElseIf types(0, k) = "HMI" Then
                If (Application.ActiveDocument.HMIObjects.Item(j).Type = types(1, k)) Then
                    ReDim Preserve objectList(1, objectListSize)
                    objectList(0, objectListSize) = Application.ActiveDocument.HMIObjects.Item(j).ObjectName
                    objectList(1, objectListSize) = Application.ActiveDocument.Name
                    objectListSize = objectListSize + 1
                End If
            End If
        Next k
    Next j
End Sub

Sub Set_Objects()
    Dim i, j, k
    
    objectsSize = 0
    ReDim objects(objectsSize)
    
    For i = 0 To ListBox_Objects.ListCount - 1
        For j = 0 To pagesSize - 1
            If ListBox_Objects.list(i, 1) = pages(j).Name Then
                For k = 1 To pages(j).HMIObjects.Count
                    If ListBox_Objects.list(i, 0) = pages(j).HMIObjects.Item(k).ObjectName Then
                        ReDim Preserve objects(objectsSize)
                        Set objects(objectsSize) = pages(j).HMIObjects.Item(k)
                        objectsSize = objectsSize + 1
                    End If
                Next k
            End If
        Next j
    Next i

    If objectsSize > 0 Then
        Lock_After_Page 3
    End If
End Sub


'--------- END Objects Functions END --------'
'=============== Objects Page ==============='

'============== Properties Page ============='
'------------ Properties Interface ----------'
Private Sub Button_Properties_SameProperties_Click()

End Sub

Private Sub Button_Properties_AllProperties_Click()
    Dim propertiesList() As String
    
    Properties_All propertiesList

    Set_ListBox ListBox_Properties, propertiesList
End Sub

Private Sub Button_Properties_EqualValue_Click()

End Sub

Private Sub Button_Properties_Clear_Click()

End Sub

Private Sub Button_Properties_ReplaceSubstring_Click()

End Sub

Private Sub Button_Properties_InsertSubstring_Click()
    Dim text
    text = InputBox("Give input", "Give Input Now", "Default" & vbCrLf & "1" & vbCrLf & "2")
    MsgBox vbCancel
    If text = "" Then
        MsgBox "Nothing entered!"
    Else
        MsgBox text
    End If
End Sub

Private Sub Button_Properties_DeleteSubstring_Click()

End Sub

Private Sub Button_Properties_SetAll_Click()
    Dim i
    
    If Not (selectEvents) Then
        For i = 0 To ListBox_Properties.ListCount - 1
            If ListBox_Properties.Selected(i) Then
                Set_All ListBox_Properties.list(i)
            End If
        Next i
    Else
    
    End If
End Sub

Private Sub ListBox_Properties_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, ListBox_Properties
End Sub

Private Sub ListBox_Events_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, ListBox_Events
End Sub

Private Sub ListBox_Properties_Change()
    Dim i
    
    ListBox_Properties.BackColor = &H80000005
    ListBox_Events.BackColor = &HC0C0C0
    selectEvents = False
    
    For i = 0 To ListBox_Properties.ListCount - 1
        If ListBox_Properties.Selected(i) Then
            Get_Properties_Value ListBox_Properties.list(i)
            
        End If
    Next i
    
End Sub

Private Sub ListBox_Options_Change()
    Dim i
    
    If Not (selectEvents) Then
        For i = 0 To ListBox_Properties.ListCount - 1
            If ListBox_Properties.Selected(i) Then
                Get_Properties_Value ListBox_Properties.list(i)
            End If
        Next i
    End If
    
End Sub


Private Sub TextBox_Value_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If valueExpanded Then
        valueExpanded = False
        
        TextBox_Value.Top = 240
        TextBox_Value.Height = 80
        TextBox_Value.ScrollBars = fmScrollBarsNone
        TextBox_Value.CurLine = 0
        
        Button_Properties_SameProperties.Visible = True
        Button_Properties_AllProperties.Visible = True
        Button_Properties_EqualValue.Visible = True
        Label_Properties.Visible = True
        Label_Events.Visible = True
        ListBox_Properties.Visible = True
        ListBox_Events.Visible = True
        Label_Options.Visible = True
        ListBox_Options.Visible = True
        Label_Dynamics.Visible = True
        ListBox_Dynamics.Visible = True
        Label_Value.Visible = True
    Else
        valueExpanded = True
        
        TextBox_Value.Top = 6
        TextBox_Value.Height = 314
        TextBox_Value.ScrollBars = fmScrollBarsBoth
        
        Button_Properties_SameProperties.Visible = False
        Button_Properties_AllProperties.Visible = False
        Button_Properties_EqualValue.Visible = False
        Label_Properties.Visible = False
        Label_Events.Visible = False
        ListBox_Properties.Visible = False
        ListBox_Events.Visible = False
        Label_Options.Visible = False
        ListBox_Options.Visible = False
        Label_Dynamics.Visible = False
        ListBox_Dynamics.Visible = False
        Label_Value.Visible = False
    End If
End Sub
'-------- END Properties Interface END ------'

'------------ Properties Functions ----------'
Sub Set_PropertiesBoth(ByRef propertiesList() As String)
    Dim propertiesListSize
    
    ListBox_Properties.Clear
    
    propertiesSize = 0
    ReDim properties(1, propertiesSize)
    
    propertiesListSize = (UBound(propertiesList, 2) - LBound(propertiesList, 2)) + 1

    For i = 0 To propertiesListSize - 1
        ReDim Preserve properties(1, propertiesSize)
        properties(0, i) = propertiesList(0, i)
        properties(1, i) = propertiesList(1, i)
        ListBox_Properties.AddItem
        ListBox_Properties.list(i) = propertiesList(0, i)
        propertiesSize = propertiesSize + 1
    Next i
End Sub

Sub Get_Properties_Value(ByVal property As String)
    Dim i, j, k
    Dim dynamics
    Dim code
    
    dynamics = 0
    code = ""
    
    For i = 0 To objectsSize - 1
        For j = 1 To objects(i).properties.Count
            If objects(i).properties.Item(j).DisplayName = property Then
                
                If ListBox_Options.Selected(0) Then
                'Static
                    
                    ListBox_Dynamics.Clear
                    
                    TextBox_Value.value = objects(i).properties.Item(j).value
                    
                ElseIf ListBox_Options.Selected(1) Then
                'Dynamic
                
                    If objects(i).properties.Item(j).IsDynamicable Then
                    
                        Set_Dynamics i, j
                        
                        dynamics = objects(i).properties.Item(j).DynamicStateType
                        If dynamics = 0 Then
                        'No dynamics
                        
                            TextBox_Value.value = ""
                            
                        ElseIf dynamics = 1 Then
                        'Direct Tag
                        
                            TextBox_Value.value = objects(i).properties.Item(j).Dynamic.VarName
                            
                        ElseIf dynamics = 2 Then
                        'Indirect Tag
                        
                            TextBox_Value.value = objects(i).properties.Item(j).Dynamic.VarName
                            
                        ElseIf dynamics = 3 Then
                        'VB or C Script
                        
                            TextBox_Value.value = objects(i).properties.Item(j).Dynamic.SourceCode
                            
                        ElseIf dynamics = 4 Then
                        'Dynamic Dialog
                        
                            code = "Formula=" & objects(i).properties.Item(j).Dynamic.SourceCode & vbLf

                            If objects(i).properties.Item(j).Dynamic.ResultType = hmiResultTypeBool Then
                            'Bool
                            
                                code = code & "True=" & objects(i).properties.Item(j).Dynamic.BinaryResultInfo.PositiveValue & vbLf
                                code = code & "False=" & objects(i).properties.Item(j).Dynamic.BinaryResultInfo.NegativeValue
                                
                            ElseIf objects(i).properties.Item(j).Dynamic.ResultType = hmiResultTypeAnalog Then
                            'Analog
                                
                                For k = 1 To objects(i).properties.Item(j).Dynamic.AnalogResultInfos.Count - 1
                                    code = code & "Value Range" & k & "="
                                    code = code & objects(i).properties.Item(j).Dynamic.AnalogResultInfos.Item(k).RangeTo & ","
                                    code = code & objects(i).properties.Item(j).Dynamic.AnalogResultInfos.Item(k).value & vbLf
                                Next k
                                
                                code = code & "Other=" & objects(i).properties.Item(j).Dynamic.AnalogResultInfos.ElseCase

                            End If

                            TextBox_Value.value = code
                            
                        End If
                    Else
                    'Not Dynamicable
                    
                        ListBox_Dynamics.Clear
                        TextBox_Value.value = ""
                        
                    End If
                    
                ElseIf ListBox_Options.Selected(2) Then
                'Update Cycle / Trigger
                
                    ListBox_Dynamics.Clear
                    
                    If objects(i).properties.Item(j).IsDynamicable Then
                    'Dynamicable
                    
                        dynamics = objects(i).properties.Item(j).DynamicStateType
                        
                        If dynamics = 1 Or dynamics = 2 Then
                        'Tag
                        
                            TextBox_Value.value = objects(i).properties.Item(j).Dynamic.CycleType
                            
                        ElseIf dynamics = 3 Or dynamics = 4 Then
                        'Script of Direct Connection
                        
                            TextBox_Value.value = objects(i).properties.Item(j).Dynamic.Trigger.Name
                            
                        Else
                        
                            TextBox_Value.value = ""
                            
                        End If
                        
                    Else
                    'Not Dynamicable
                    
                        TextBox_Value.value = ""
                        
                    End If
                ElseIf ListBox_Options.Selected(3) Then
                'Indirect
                
                    ListBox_Dynamics.Clear
                    
                    If objects(i).properties.Item(j).IsDynamicable Then
                    'Dynamicable
                    
                        If objects(i).properties.Item(j).DynamicStateType = 1 Then
                        'Direct Tag
                        
                            TextBox_Value.value = "0"
                            
                        ElseIf objects(i).properties.Item(j).DynamicStateType = 2 Then
                        'Indirect Tag
                        
                            TextBox_Value.value = "1"
                            
                        Else
                        'Neither
                        
                            TextBox_Value.value = ""
                            
                        End If
                        
                    Else
                    'Not dynamicable
                    
                        TextBox_Value.value = ""
                        
                    End If
                End If
                
                'Break out of loops when found
                j = objects(i).properties.Count + 1
                i = objectsSize
            End If
        Next j
    Next i
End Sub

Sub Set_Dynamics(ByVal objNum As Integer, ByVal propNum As Integer)
    Dim dynamics, codeType
    Dim dynamicsList(3) As String
    dynamicsList(0) = "Dynamic Dialog"
    dynamicsList(1) = "C-Script"
    dynamicsList(2) = "VBScript"
    dynamicsList(3) = "Tag"
    
    Set_ListBox ListBox_Dynamics, dynamicsList
    
    dynamics = objects(objNum).properties.Item(propNum).DynamicStateType
    
    If dynamics = 0 Then
    'No dynamics

    ElseIf dynamics = 1 Then
    'Direct Tag
        ListBox_Dynamics.Selected(3) = True
    ElseIf dynamics = 2 Then
    'Indirect Tag
        ListBox_Dynamics.Selected(3) = True
    ElseIf dynamics = 3 Then
    'VB or C Script
        If objects(objNum).properties.Item(propNum).Dynamic.ScriptType = 0 Then
        '0 = VBS
            ListBox_Dynamics.Selected(2) = True
        Else
        '1 = C
            ListBox_Dynamics.Selected(1) = True
        End If
    ElseIf dynamics = 4 Then
    'Direct connection
        ListBox_Dynamics.Selected(0) = True
    End If
End Sub

Sub Set_All(ByRef property As String)
    Dim i, j, k
    Dim dynamics
    Dim code
    
    Dim dynDialog As HMIDynamicDialog
    Dim parseOk, text
    Dim formula As String, posVal As Long, negVal As Long, rangeVals() As Long, rangeValsSize As Long, elseVal As Long
    rangeValsSize = 0
    
    dynamics = 0
    code = ""
    
    For i = 0 To objectsSize - 1
        For j = 1 To objects(i).properties.Count
            If objects(i).properties.Item(j).DisplayName = property Then
                
                If ListBox_Options.Selected(0) Then
                'Static
                    
                    objects(i).properties.Item(j).value = TextBox_Value.value
                    
                ElseIf ListBox_Options.Selected(1) Then
                'Dynamic
                    If objects(i).properties.Item(j).IsDynamicable Then
                        
                        objects(i).properties.Item(j).DeleteDynamic
                        
                        If ListBox_Dynamics.Selected(0) Then
                        'Dynamic Dialog
                            
                            'For Parsing Errors
                            'On Error Resume Next
                            
                            parseOk = True
                            text = TextBox_Value.value
                            
                            If InStr(text, "Formula=") = 1 Then
                                text = Replace(text, "Formula=", "")
                            Else
                                parseOk = False
                            End If

                            formula = Split(text, vbLf, 2)(0)
                            text = Split(text, vbLf, 2)(1)

                            If InStr(text, "True=") = 1 Then
                                
                                text = Replace(text, "True=", "")
                                MsgBox Split(text, vbLf, 2)(0)
                                posVal = Split(text, vbLf, 2)(0)
                                text = Split(text, vbLf, 2)(1)
                                
                                If InStr(text, "False=") = 1 Then
                                    text = Replace(text, "False=", "")
                                    negVal = Split(text, vbLf, 2)(0)
                                Else
                                    parseOk = False
                                End If
                                
                            ElseIf InStr(text, "Value Range") = 1 Or InStr(text, "Other") = 1 Then

                                rangeValsSize = 0
                                ReDim rangeVals(1, rangeValsSize)
                                
                                While InStr(text, "Value Range") = 1
                                    ReDim Preserve rangeVals(1, rangeValsSize)

                                    text = Replace(text, "Value Range" & rangeValsSize + 1 & "=", "")

                                    rangeVals(0, rangeValsSize) = Split(text, ",", 2)(0)
                                    text = Split(text, ",", 2)(1)

                                    rangeVals(1, rangeValsSize) = Split(text, vbLf, 2)(0)
                                    text = Split(text, vbLf, 2)(1)
                                    
                                    rangeValsSize = rangeValsSize + 1
                                Wend
                                
                                If InStr(text, "Other=") = 1 Then
                                    text = Replace(text, "Other=", "")
                                    elseVal = text
                                Else
                                    parseOk = False
                                End If
                                
                                If parseOk Then
                                    
                                    Set dynDialog = objects(i).properties.Item(j).CreateDynamic(hmiDynamicCreationTypeDynamicDialog, formula)
                                    dynDialog.ResultType = hmiResultTypeAnalog
                                    For k = 0 To rangeValsSize - 1
                                        dynDialog.AnalogResultInfos.Add rangeVals(0, k), rangeVals(1, k)
                                        'MsgBox rangeVals(0, k) & "," & rangeVals(1, k)
                                    Next k
                                    dynDialog.AnalogResultInfos.ElseCase = elseVal
                                End If
                            Else
  
                            End If
                            
                            If Err.Number <> 0 Then
                                objects(i).properties.Item(j).DeleteDynamic
                                Err.Clear
                            End If
                            
                        ElseIf ListBox_Dynamics.Selected(1) Then
                        'C-Script
                            
                            objects(i).properties.Item(j).CreateDynamic hmiDynamicCreationTypeCScript
                            objects(i).properties.Item(j).Dynamic.SourceCode = TextBox_Value.value
                            
                        ElseIf ListBox_Dynamics.Selected(2) Then
                        'VBScript
                            
                            objects(i).properties.Item(j).CreateDynamic hmiDynamicCreationTypeVBScript
                            objects(i).properties.Item(j).Dynamic.SourceCode = TextBox_Value.value
                            
                        ElseIf ListBox_Dynamics.Selected(3) Then
                        'Tag
                        
                            objects(i).properties.Item(j).CreateDynamic hmiDynamicCreationTypeVariableDirect, TextBox_Value.value
                            objects(i).properties.Item(j).Dynamic.CycleType = 11

                        End If
                        
                    End If
                    
                ElseIf ListBox_Options.Selected(2) Then
                'Update Cycle / Trigger
                    
                    If objects(i).properties.Item(j).IsDynamicable Then
                    
                        dynamics = objects(i).properties.Item(j).DynamicStateType
                        
                        If dynamics = 1 Or dynamics = 2 Then
                        'Tag
                            objects(i).properties.Item(j).Dynamic.CycleType = TextBox_Value.value
                        ElseIf dynamics = 3 Or dynamics = 4 Then
                        'Script or Direct Connection
                            objects(i).properties.Item(j).Dynamic.Trigger.Name = TextBox_Value.value
                        End If
                        
                    Else
                        TextBox_Value.value = ""
                    End If
                    
                ElseIf ListBox_Options.Selected(3) Then
                'Indirect
                
                    ListBox_Dynamics.Clear
                    
                    If objects(i).properties.Item(j).IsDynamicable Then
                        If (TextBox_Value.value = 0) Or (UCase(TextBox_Value.value) = "FALSE") Then
                            'objects(i).properties.Item(j).DynamicStateType = 1
                        ElseIf (TextBox_Value.value = 1) Or (UCase(TextBox_Value.value) = "TRUE") Then
                            'objects(i).properties.Item(j).DynamicStateType = 2
                        End If
                    End If
                End If
                
            End If
        Next j
    Next i
End Sub

Sub Properties_Same(ByRef propertiesList() As String)

End Sub

Sub Properties_All(ByRef propertiesList() As String)
    Dim i, j, m, n, p, q
    Dim propertiesListSize
    Dim propertyIncluded
    
    propertiesListSize = 0
    ReDim propertiesList(propertiesListSize)
    
    For i = 0 To objectsSize - 1
        For j = 1 To objects(i).properties.Count
            propertyIncluded = False
            For k = 0 To propertiesListSize - 1
                If propertiesList(k) = objects(i).properties.Item(j).DisplayName Then
                    propertyIncluded = True
                End If
            Next k
            
            If Not (propertyIncluded) Then
                ReDim Preserve propertiesList(propertiesListSize)
                propertiesList(propertiesListSize) = objects(i).properties.Item(j).DisplayName
                propertiesListSize = propertiesListSize + 1
            End If
        Next j
    Next i
End Sub

Sub Properties_Equal(ByRef propertiesList() As String)

End Sub

Function Properties_FirstValue()
    Dim i, j, k
End Function
'-------- END Properties Functions END ------'
'========== END Properties Page END ========='

