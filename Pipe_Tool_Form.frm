VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pipe_Tool_Form 
   Caption         =   "Pipe Tool"
   ClientHeight    =   5775
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   3915
   OleObjectBlob   =   "Pipe_Tool_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pipe_Tool_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pipes() As HMITubePolyline
Dim pipesSize As Integer
Dim lastPipeSetForDivide As String
Dim divideChangeFromSub As Boolean

Private Declare Function MessageBox _
    Lib "user32" Alias "MessageBoxA" _
       (ByVal hWnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal wType As Long) _
    As Long

Private Sub Frame_Animation_Click()
    
End Sub

Private Sub Button_Object_Connect_Click()

End Sub

Private Sub Button_Test_Click()

End Sub

' ================ Form Setup ================
Private Sub UserForm_Activate()
'Setup upon opening form
    
    'Disable until implementation
    Button_Extend_Pipes.Enabled = False
    
    Set_Pipes_Selected
    
    If pipesSize = 1 Then
    'Set fields for division
    
        Pipe_Tool_Form.TextBox_Min.value = "1"
        Pipe_Tool_Form.TextBox_Max.value = Int(Get_Pipe_Length(pipes(0)) - 1)
        Pipe_Tool_Form.TextBox_Divide.value = Int(Get_Pipe_Length(pipes(0)) / 2)
        
    End If 'pipesSize = 1
    
End Sub

' ================ Merge Interface ================
Private Sub Button_Merge_Pipes_Click()
    Dim closestPipes() As Integer
    
    Set_Pipes_Selected
    
    If pipesSize > 1 Then
    'Check for correct number of pipes
    
        While pipesSize > 1
            closestPipes = Get_Closest_Pipes(pipes)
            Merge_Pipes pipes(closestPipes(0)), pipes(closestPipes(1))
        Wend
        
        Select_Pipes
    Else 'pipesSize > 1
        
        MsgBox "Two or more pipes must be selected for merging", vbExclamation, "Selection Error"
        
    End If 'pipesSize > 1

End Sub

' ================ Split Interface ================
Private Sub Button_Split_Pipe_Click()
    Dim i
    
    Set_Pipes_Selected
    
    If pipesSize > 0 Then
    'Check that at least one pipe is selected
       
        For i = 0 To pipesSize - 1

            If pipes(i).PointCount > 2 Then
            
                Split_Pipe pipes(i)
                i = i - 1
                
            End If
            
        Next i
    
        Select_Pipes
        
    Else
    
        MsgBox "Must have at least one pipe selected for splitting.", vbExclamation, "Selection Error"
    
    End If
    
    Select_Pipes
End Sub

' ================ Extend Interface ================
Private Sub Button_Extend_Pipes_Click()
    'Dim closestPipes() As Integer
    
    'Set_Pipes_Selected
    
    'If pipesSize > 1 Then
    'Check for correct number of pipes
    
        'While pipesToExtendSize
            'closestPipes = Get_Closest_Pipes(pipes)
            'Merge_Pipes pipes(closestPipes(0)), pipes(closestPipes(1))
        'Wend
        
        'Select_Pipes
    'Else 'pipesSize > 1
        
        'MsgBox "Two or more pipes must be selected for merging", vbExclamation, "Selection Error"
        
    'End If 'pipesSize > 1

End Sub

' ================ Remove Middle Nodes Interface ================
Private Sub Button_Remove_Mid_Nodes_Click()
    Dim i
    
    Set_Pipes_Selected
    
    If pipesSize > 0 Then
    'Check that at least one pipe is selected
        
        For i = 0 To pipesSize - 1
        
            Remove_Mid_Nodes pipes(i)
            
        Next i
    
    Else
    
        MsgBox "Must have at least one pipe selected for removing middle nodes.", vbExclamation, "Selection Error"
    
    End If
End Sub

' ================ Allign Interface ================
Private Sub Button_Allign_Click()
    Dim i
    Dim pipeReference As HMITubePolyline
    Dim pipeRefSet
    
    Set_Pipes_Selected
    
    If pipesSize > 1 Then
    'Check for correct number of pipes
        Unselect_Pipes
        
        MessageBox &O0, "Select reference pipe then press ok", "Select Reference Pipe", vbOKOnly + vbSystemModal
        
        pipeRefSet = False
        
        For i = 1 To ActiveDocument.HMIObjects.Count
            If ActiveDocument.HMIObjects.Item(i).Selected Then
                If ActiveDocument.HMIObjects.Item(i).Type = "HMITubePolyline" Then
                    pipeRefSet = True
                    Set pipeReference = ActiveDocument.HMIObjects.Item(i)
                End If
            End If
        Next i
        
        If pipeRefSet Then
        
            For i = 0 To pipesSize - 1
                
                If pipeReference.ObjectName <> pipes(i).ObjectName Then
                'Don't allign pipe to itself
                
                    Allign_Pipe pipeReference, pipes(i)
                
                End If
                
            Next i
            
        Else
            
            MsgBox "No reference pipe selected", vbExclamation, "Selection Error"
            
        End If
        
    Else 'pipesSize > 1
    
        MsgBox "Two or more pipes must be selected for alligning", vbExclamation, "Selection Error"
    
    End If 'pipesSize > 1
    
    Select_Pipes
    
End Sub

' ================ Straighet Interface ================
Private Sub Button_Straighten_Pipes_Click()
    Dim i
    
    Set_Pipes_Selected
    
    If pipesSize > 0 Then
    'Ensure at least one pipe is selected
    
        For i = 0 To pipesSize - 1
            Straighten_Pipe pipes(i)
        Next i
        
        Select_Pipes
        
    Else 'pipesSize > 0
    
        MsgBox "Must have at least one pipe selected for straightening.", vbExclamation, "Selection Error"
    
    End If 'pipesSize > 0
    
End Sub

' ================ Divide Interface ================
Private Sub Button_Divide_Pipe_Click()
    Set_Pipes_Selected
    
    If pipesSize = 1 Then
    'Ensure correct number of pipes for operation
        
        TextBox_Min.value = "1"
        TextBox_Max.value = Int(Get_Pipe_Length(pipes(0)) - 1)
        
        If TextBox_Divide.value < 1 Or TextBox_Divide.value > TextBox_Max.value Then
            TextBox_Divide.value = Int(Get_Pipe_Length(pipes(0)) / 2)
        End If
        
        If Get_Pipe_Length(pipes(0)) > 1 Then
        'Ensure pipe is long enough to be divided
        
            If pipes(0).PointCount = 2 Then
            'Ensure correct number of vertices
            
                Divide_Pipe pipes(0), Int(TextBox_Divide.value)
                Select_Pipes
        
            Else 'pipes(0).PointCount = 2
            
                MsgBox "Only pipes with 2  vertices can be selected for division.", vbExclamation, "Selection Error"
            
            End If 'pipes(0).PointCount = 2
        
        Else 'Get_Pipe_Length(pipes(0)) > 1
        
            MsgBox "Only pipes 2 units or longer can be selected for division.", vbExclamation, "Selection Error"
        
        End If 'Get_Pipe_Length(pipes(0)) > 1
        
    Else 'pipesSize = 1 Then
    
        MsgBox "Only one pipe can be selected for division.", vbExclamation, "Selection Error"
        
    End If 'pipesSize = 1 Then
    
End Sub

Private Sub TextBox_Divide_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set_Pipes_Selected
    
    If pipesSize = 1 Then
        If pipes(0).PointCount = 2 Then
        
        'Check for correct number of pipes
            TextBox_Min.value = "1"
            TextBox_Max.value = Int(Get_Pipe_Length(pipes(0)) - 1)
            
            If lastPipeSetForDivide = "" Then
            'If no pipe was set
                lastPipeSetForDivide = pipes(0).ObjectName
                divideChangeFromSub = True
                TextBox_Divide.value = Int(Get_Pipe_Length(pipes(0)) / 2)
            End If
            
            If pipes(0).ObjectName <> lastPipeSetForDivide Then
            'If pipe changed reset to middle
                divideChangeFromSub = True
                TextBox_Divide.value = Int(Get_Pipe_Length(pipes(0)) / 2)
                lastPipeSetForDivide = pipes(0).ObjectName
            End If
        End If
    End If 'pipesSize = 1
    
End Sub

Private Sub TextBox_Divide_Change()
'Input validation for divide text box
    If Not (divideChangeFromSub) Then

        If pipesSize = 1 Then
            If pipes(0).PointCount = 2 Then
                
                lastPipeSetForDivide = pipes(0).ObjectName
                
                If IsNumeric(TextBox_Divide.value) Then
                'If value is a number continue check
                    
                    If Int(TextBox_Divide.value) < 1 Then
                    'If value is less than min, set to min
                    
                        TextBox_Divide.value = 1
                        
                    End If 'Int(TextBox_Divide.value) < 1
                    
                    If Int(TextBox_Divide.value) > Get_Pipe_Length(pipes(0)) - 1 Then
                    'If value is more than max set to max
                    
                        TextBox_Divide.value = Get_Pipe_Length(pipes(0)) - 1
                        
                    End If 'Int(TextBox_Divide.value) > Get_Pipe_Length(pipes(0)) - 1
                    
                Else 'IsNumeric(TextBox_Divide.value)
                'If number is not a value default to 1
                
                    TextBox_Divide.value = 1
                    
                End If 'IsNumeric(TextBox_Divide.value)
            End If
        Else
            TextBox_Min.value = ""
            TextBox_Max.value = ""
            TextBox_Divide.value = ""
        End If
    Else
        divideChangeFromSub = False
    End If
End Sub

' ================ Add Connectors Interface ================
Private Sub Button_Add_Connectors_Click()
    Dim i, j, k
    Dim allCoords() '(0,x) = x coord, (1,x) = y coord, (2,x) = connection directions
    Dim allCoordsSize
    Dim addNewCoord
    Dim coord(1) As Integer
    Dim connectMiddleNodes
    
    connectMiddleNodes = False
    
    allCoordsSize = 0
    
    'Directions
    '1 = Up
    '2 = Right
    '4 = Down
    '8 = Left
    'E.G 3 = Up + Right, 11 = Up + Right + Left, 15 = Up + Right + Down + Left
    'If only 1, 2, 4, or 8 no connector will be made
    
    Set_Pipes_Selected
    
    For i = 0 To pipesSize - 1
        For j = 1 To pipes(i).PointCount
            If (j = 1 Or j = pipes(i).PointCount) Or connectMiddleNodes Then
                addNewCoord = True
                pipes(i).index = j
                
                For k = 0 To allCoordsSize - 1
                    If (allCoords(0, k) = pipes(i).ActualPointLeft) And (allCoords(1, k) = pipes(i).ActualPointTop) Then
                        addNewCoord = False
                        allCoords(2, k) = allCoords(2, k) + Pipe_Connector_Direction(pipes(i), j)
                        k = allCoordsSize 'go ahead and break loop
                    End If
                Next k
                
                If addNewCoord Then
                    ReDim Preserve allCoords(2, allCoordsSize)
                    allCoords(0, allCoordsSize) = pipes(i).ActualPointLeft
                    allCoords(1, allCoordsSize) = pipes(i).ActualPointTop
                    allCoords(2, allCoordsSize) = Pipe_Connector_Direction(pipes(i), j)
                    allCoordsSize = allCoordsSize + 1
                End If
            End If
        Next j
    Next i
    
    For i = 0 To allCoordsSize - 1
        coord(0) = allCoords(0, i)
        coord(1) = allCoords(1, i)
        Add_Connector coord, allCoords(2, i)
    Next i
    
End Sub

' ================ Main Functions ================
Sub Split_Pipe(ByVal pipe As HMITubePolyline)
    Dim i
    Dim oldCoords() As Integer
    Dim newCoords(1, 1) As Integer
    Dim numCoords As Integer
    Dim newPipeName As String
    
    oldCoords = Get_Coords(pipe)
    numCoords = UBound(oldCoords, 2) - LBound(oldCoords, 2) + 1
    
    Dim pipeNum
    For i = 0 To pipesSize - 1
        If pipes(i).ObjectName = pipe.ObjectName Then
            pipeNum = i
        End If
    Next i
    
    Delete_Pipe pipe
    
    For i = 0 To numCoords - 2
        newCoords(0, 0) = oldCoords(0, i)
        newCoords(1, 0) = oldCoords(1, i)
        newCoords(0, 1) = oldCoords(0, i + 1)
        newCoords(1, 1) = oldCoords(1, i + 1)
        Create_Pipe newCoords
    Next i
    
    'ActiveDocument.HMIObjects(pipeName).Delete
End Sub

Sub Divide_Pipe(ByVal pipe As HMITubePolyline, divPoint As Integer)
    Dim oldCoords() As Integer
    Dim middleCoords(1, 0) As Integer
    Dim newCoords(1, 1) As Integer
    Dim numCoords As Integer
    Dim fraction As Double
    
    oldCoords = Get_Coords(pipe)
    numCoords = UBound(oldCoords, 2) - LBound(oldCoords, 2) + 1

    If numCoords = 2 Then
    
        If divPoint > 1 And divPoint < Get_Pipe_Length(pipe) Then
        
            fraction = divPoint / Get_Pipe_Length(pipe)
            middleCoords(0, 0) = Int((Get_Pipe_X_Length(pipe) * fraction) + Get_Pipe_X_Min(pipe))
            middleCoords(1, 0) = Int((Get_Pipe_Y_Length(pipe) * fraction) + Get_Pipe_Y_Min(pipe))
            
            Delete_Pipe pipe
            
            newCoords(0, 0) = oldCoords(0, 0)
            newCoords(1, 0) = oldCoords(1, 0)
            newCoords(0, 1) = middleCoords(0, 0)
            newCoords(1, 1) = middleCoords(1, 0)
            
            Create_Pipe newCoords
            
            newCoords(0, 0) = middleCoords(0, 0)
            newCoords(1, 0) = middleCoords(1, 0)
            newCoords(0, 1) = oldCoords(0, 1)
            newCoords(1, 1) = oldCoords(1, 1)
            
            Create_Pipe newCoords
            
        End If 'divPoint > 1 And divPoint < Get_Pipe_Length(pipe)
        
    End If 'numCoords = 2
    
End Sub

Sub Merge_Pipes(ByVal pipe1 As HMITubePolyline, ByVal pipe2 As HMITubePolyline)
    Dim i
    Dim newCoords() As Integer
    Dim newCoordsSize As Integer
    Dim coordIntersect() As Integer
    Dim closestVerts() As Integer
    
    closestVerts = Get_Closest_Vertices(pipe1, pipe2)
    
    Pipe_Parallel_Allign pipe1, closestVerts(0), pipe2, closestVerts(1)
    
    newCoordsSize = pipe1.PointCount + pipe2.PointCount - 2
    ReDim newCoords(1, newCoordsSize)
    
    coordIntersect = Pipe_Intersect(pipe1, closestVerts(0), pipe2, closestVerts(1))
    
    For i = 0 To pipe1.PointCount - 2
    
        If closestVerts(0) = 1 Then
        'Start at the back of the pipe and skip point 1
            pipe1.index = pipe1.PointCount - i
        Else
        'Start at the front of the pipe and skip the last point
            pipe1.index = i + 1
        End If
        
        newCoords(0, i) = pipe1.ActualPointLeft
        newCoords(1, i) = pipe1.ActualPointTop
        
    Next i
    
    newCoords(0, pipe1.PointCount - 1) = coordIntersect(0)
    newCoords(1, pipe1.PointCount - 1) = coordIntersect(1)
    
    For i = pipe1.PointCount To newCoordsSize
        
        If closestVerts(1) = 1 Then
        'Start at the front of the pipe skipping the first point
            pipe2.index = i + 2 - pipe1.PointCount
        Else
        'Start at the back of the pipe skipping the last point
            pipe2.index = newCoordsSize - i + 1
        End If

        newCoords(0, i) = pipe2.ActualPointLeft
        newCoords(1, i) = pipe2.ActualPointTop

    Next i
    
    Delete_Pipe pipe1
    Delete_Pipe pipe2

    Create_Pipe newCoords

End Sub

Sub Allign_Pipe(ByVal pipeReference As HMITubePolyline, ByVal pipeMove As HMITubePolyline)
    Dim i
    Dim closestVerts() As Integer
    Dim xDelta, yDelta
    
    closestVerts = Get_Closest_Vertices(pipeReference, pipeMove)
    
    pipeReference.index = closestVerts(0)
    pipeMove.index = closestVerts(1)
    
    xDelta = pipeReference.ActualPointLeft - pipeMove.ActualPointLeft
    yDelta = pipeReference.ActualPointTop - pipeMove.ActualPointTop
    
    pipeMove.Left = pipeMove.Left + xDelta
    pipeMove.Top = pipeMove.Top + yDelta

End Sub

Sub Remove_Mid_Nodes(ByVal pipe As HMITubePolyline)
    Dim i, j
    Dim x1, x2, x3, y1, y2, y3
    
    If pipe.PointCount > 2 Then
        i = 2
        j = 0
        While i < pipe.PointCount
        'For i = 2 To (pipe.PointCount - 1)
            pipe.index = i - 1
            x1 = pipe.ActualPointLeft
            y1 = pipe.ActualPointTop
            pipe.index = i
            x2 = pipe.ActualPointLeft
            y2 = pipe.ActualPointTop
            pipe.index = i + 1
            x3 = pipe.ActualPointLeft
            y3 = pipe.ActualPointTop
            
            If ((x1 = x2) And (x2 = x2) And (x3 = x1)) Or ((y1 = y2) And (y2 = y2) And (y3 = y1)) Then
                Remove_Vertex pipe, i
                j = j + 1
            Else
                i = i + 1
                j = 0
            End If
            
            If j > 1000 Then
            'Inifite for loop check
                MsgBox "Error ininite loop detected in sub 'Remove_Mid_Nodes'. Exceeded 1000 loops. Manually quiting.", vbOKOnly, "Error in sub!"
            End If
            
        Wend

    End If
    
End Sub

Sub Straighten_Pipe(ByVal pipe As HMITubePolyline)
    Dim i
    Dim x1, x2, y1, y2, xLength, yLength, xCenter, yCenter, lineLength
    Dim firstLine
    
    firstLine = True
    
    For i = 1 To pipe.PointCount - 1
        lineLength = Vertex_Distance(pipe, i, pipe, i + 1)
        pipe.index = i
        x1 = pipe.ActualPointLeft
        y1 = pipe.ActualPointTop
        
        pipe.index = i + 1
        x2 = pipe.ActualPointLeft
        y2 = pipe.ActualPointTop
        
        xCenter = x1 + ((x2 - x1) / 2)
        yCenter = y1 + ((y2 - y1) / 2)
        
        xLength = Abs(x1 - x2)
        yLength = Abs(y1 - y2)
        
        If yLength > xLength Then
            If y1 > y2 Then
                If firstLine Then
                    pipe.index = i
                    pipe.ActualPointLeft = xCenter
                    pipe.ActualPointTop = yCenter + (lineLength / 2)
                Else
                    pipe.index = i
                    xCenter = pipe.ActualPointLeft
                End If
                
                pipe.index = i + 1
                pipe.ActualPointLeft = xCenter
                pipe.ActualPointTop = yCenter - (lineLength / 2)
            Else
                If firstLine Then
                    pipe.index = i
                    pipe.ActualPointLeft = xCenter
                    pipe.ActualPointTop = yCenter - (lineLength / 2)
                Else
                    pipe.index = i
                    xCenter = pipe.ActualPointLeft
                End If
                
                pipe.index = i + 1
                pipe.ActualPointLeft = xCenter
                pipe.ActualPointTop = yCenter + (lineLength / 2)
            End If

            firstLine = False
        Else
            If x1 > x2 Then
                If firstLine Then
                    pipe.index = i
                    pipe.ActualPointLeft = xCenter + (lineLength / 2)
                    pipe.ActualPointTop = yCenter
                Else
                    pipe.index = i
                    yCenter = pipe.ActualPointTop
                End If
                
                pipe.index = i + 1
                pipe.ActualPointLeft = xCenter - (lineLength / 2)
                pipe.ActualPointTop = yCenter
            Else
                If firstLine Then
                    pipe.index = i
                    pipe.ActualPointLeft = xCenter - (lineLength / 2)
                    pipe.ActualPointTop = yCenter
                Else
                    pipe.index = i
                    yCenter = pipe.ActualPointTop
                End If
                
                pipe.index = i + 1
                pipe.ActualPointLeft = xCenter + (lineLength / 2)
                pipe.ActualPointTop = yCenter
            End If
            
            firstLine = False
        End If
    Next i
End Sub

' ================ Add_Connector ================
Sub Add_Connector(ByRef coord() As Integer, ByVal direction As Integer)
    Dim newPipe As HMITubePolyline
    Dim newTee As HMITubeTeeObject
    Dim newDoubleTee As HMITubeDoubleTeeObject
    
    If direction = 5 Then
    'Up + Down
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0)
        newPipe.Top = coord(1) - 10
        
        newPipe.PointCount = 2
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) - 10
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) + 10
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 0
        newPipe.Height = 20
    ElseIf direction = 10 Then
    'Left + Right
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0) - 10
        newPipe.Top = coord(1)
        
        newPipe.PointCount = 2
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0) - 10
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0) + 10
        newPipe.ActualPointTop = coord(1)
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 20
        newPipe.Height = 0
    ElseIf direction = 3 Then
    'Up + Right
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0)
        newPipe.Top = coord(1) - 10
        
        newPipe.PointCount = 3
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) - 10
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 3
        newPipe.ActualPointLeft = coord(0) + 10
        newPipe.ActualPointTop = coord(1)
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 10
        newPipe.Height = 10
    ElseIf direction = 6 Then
    'Right + Down
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0)
        newPipe.Top = coord(1)
        
        newPipe.PointCount = 3
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0) + 10
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 3
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) + 10
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 10
        newPipe.Height = 10
    ElseIf direction = 12 Then
    'Down + Left
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0) - 10
        newPipe.Top = coord(1)
        
        newPipe.PointCount = 3
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) + 10
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 3
        newPipe.ActualPointLeft = coord(0) - 10
        newPipe.ActualPointTop = coord(1)
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 10
        newPipe.Height = 10
    ElseIf direction = 9 Then
    'Left + Up
        Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon Tube", "HMITubePolyline")
        newPipe.Layer = 1
        newPipe.Left = coord(0) - 10
        newPipe.Top = coord(1) - 10
        
        newPipe.PointCount = 3
        newPipe.index = 1
        newPipe.ActualPointLeft = coord(0) - 10
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 2
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1)
        newPipe.index = 3
        newPipe.ActualPointLeft = coord(0)
        newPipe.ActualPointTop = coord(1) - 10
        
        newPipe.GlobalColorScheme = False
        newPipe.BorderColor = RGB(182, 182, 182)
        newPipe.BorderWidth = 15
        newPipe.Width = 10
        newPipe.Height = 10
    ElseIf direction = 11 Then
    'Left + Up + Right
        Set newTee = ActiveDocument.HMIObjects.AddHMIObject("T-piece", "HMITubeTeeObject")
        newTee.Layer = 1
        newTee.Left = coord(0) - 10
        newTee.Top = coord(1) - 10
        
        newTee.GlobalColorScheme = False
        newTee.BorderColor = RGB(182, 182, 182)
        newTee.BorderWidth = 15
        newTee.Width = 20
        newTee.Height = 10
        newTee.RotationAngle = 180
    ElseIf direction = 7 Then
    'Up + Right + Down
        Set newTee = ActiveDocument.HMIObjects.AddHMIObject("T-piece", "HMITubeTeeObject")
        newTee.Layer = 1
        newTee.Left = coord(0)
        newTee.Top = coord(1) - 10
        
        newTee.GlobalColorScheme = False
        newTee.BorderColor = RGB(182, 182, 182)
        newTee.BorderWidth = 15
        newTee.Width = 10
        newTee.Height = 20
        newTee.RotationAngle = 270
    ElseIf direction = 14 Then
    'Right + Down + Left
        Set newTee = ActiveDocument.HMIObjects.AddHMIObject("T-piece", "HMITubeTeeObject")
        newTee.Layer = 1
        newTee.Left = coord(0) - 10
        newTee.Top = coord(1)
        
        newTee.GlobalColorScheme = False
        newTee.BorderColor = RGB(182, 182, 182)
        newTee.BorderWidth = 15
        newTee.Width = 20
        newTee.Height = 10
        newTee.RotationAngle = 0
    ElseIf direction = 13 Then
    'Down + Left + Up
        Set newTee = ActiveDocument.HMIObjects.AddHMIObject("T-piece", "HMITubeTeeObject")
        newTee.Layer = 1
        newTee.Left = coord(0) - 10
        newTee.Top = coord(1) - 10
        
        newTee.GlobalColorScheme = False
        newTee.BorderColor = RGB(182, 182, 182)
        newTee.BorderWidth = 15
        newTee.Width = 10
        newTee.Height = 20
        newTee.RotationAngle = 90
    ElseIf direction = 15 Then
    'All
        Set newDoubleTee = ActiveDocument.HMIObjects.AddHMIObject("Double T-piece", "HMITubeDoubleTeeObject")
        newDoubleTee.Layer = 1
        newDoubleTee.Left = coord(0) - 10
        newDoubleTee.Top = coord(1) - 10
        
        newDoubleTee.GlobalColorScheme = False
        newDoubleTee.BorderColor = RGB(182, 182, 182)
        newDoubleTee.BorderWidth = 15
        newDoubleTee.Width = 20
        newDoubleTee.Height = 20
    End If
    
End Sub

Sub Object_Connect(ByRef obj As HMIObject)
    If obj.Type = "HMIFaceplateObject" And obj.properties.Item(3) = "fpWaterWellPump.fpt" Then
    
    ElseIf obj.Type = "HMIFaceplateObject" And obj.properties.Item(3) = "fpWaterWellPump.fpt" Then
    
    Else
    
    End If
    
End Sub



' ================ Auxiliary Functions ================
Sub Select_Pipes()
    Dim i
    
    For i = 0 To pipesSize - 1
        pipes(i).Selected = True
    Next i
End Sub

Sub Unselect_Pipes()
    Dim i
    
    For i = 0 To pipesSize - 1
        pipes(i).Selected = False
    Next i
End Sub

Sub Set_Pipes_Selected()
    Dim i
    
    pipesSize = 0
    ReDim pipes(pipesSize)
    
    For i = 1 To ActiveDocument.HMIObjects.Count
        If ActiveDocument.HMIObjects.Item(i).Selected Then
            If ActiveDocument.HMIObjects.Item(i).Type = "HMITubePolyline" Then
            'Ensure object is an HMITubePolyline
                If ActiveDocument.HMIObjects.Item(i).BorderWidth = 10 Then
                'Pipes of 15 BorderWidth are used as connectors and are not added
                    ReDim Preserve pipes(pipesSize)
                    Set pipes(pipesSize) = ActiveDocument.HMIObjects.Item(i)
                    pipesSize = pipesSize + 1
                End If
            End If
        End If
    Next i
       
End Sub

Sub Create_Pipe(coords() As Integer)
    Dim i, numCoords
    Dim newPipe As HMITubePolyline
    Set newPipe = ActiveDocument.HMIObjects.AddHMIObject("Polygon tube", "HMITubePolyline")
    
    numCoords = UBound(coords, 2) - LBound(coords, 2) + 1
    
    newPipe.GlobalColorScheme = False
    newPipe.BorderColor = RGB(182, 182, 182)
    newPipe.BorderWidth = 10
    newPipe.PointCount = numCoords
    
    For i = 0 To numCoords - 1
        newPipe.index = i + 1
        newPipe.ActualPointLeft = coords(0, i)
        newPipe.ActualPointTop = coords(1, i)
    Next i
    
    newPipe.Left = Get_Pipe_X_Min(newPipe)
    newPipe.Top = Get_Pipe_Y_Min(newPipe)
    
    ReDim Preserve pipes(pipesSize)
    Set pipes(pipesSize) = newPipe
    pipesSize = pipesSize + 1
    
End Sub

Sub Delete_Pipe(ByVal pipe As HMITubePolyline)
    Dim i, j
    Dim newPipes() As HMITubePolyline
    
    If pipesSize > 1 Then
    'If more than one pipe exists

        ReDim newPipes(pipesSize - 2)
        j = 0
        For i = 0 To pipesSize - 1
            If pipes(i).ObjectName <> pipe.ObjectName Then
            'If this is not the pipe to delete
                
             Set newPipes(j) = pipes(i)
             j = j + 1
            
            End If 'i <> pipeNum
        Next i
        
        pipe.Delete
        ReDim pipes(pipesSize - 2)
        
        For i = 0 To pipesSize - 2
            Set pipes(i) = newPipes(i)
        Next i
        
        pipesSize = pipesSize - 1
        
    Else 'pipesSize > 1
    'If only one pipe exists simply clear pipes()
    
        pipe.Delete
        ReDim pipes(0)
        pipesSize = 0
    
    End If 'pipesSize > 1
End Sub

Function Get_Coords(ByVal pipe As HMITubePolyline)
    Dim i, coordsSize
    Dim coords() As Integer
    
    ReDim coords(1, 0)
    coordsSize = 0
    
    For i = 1 To pipe.PointCount
        pipe.index = i
        ReDim Preserve coords(1, coordsSize)
        coords(0, i - 1) = pipe.ActualPointLeft
        coords(1, i - 1) = pipe.ActualPointTop
        coordsSize = coordsSize + 1
    Next i
    
    Get_Coords = coords
    
End Function

Function Get_Pipe_Length(ByVal pipe As HMITubePolyline)
    Get_Pipe_Length = Sqr((Get_Pipe_Y_Length(pipe)) ^ 2 + (Get_Pipe_X_Length(pipe)) ^ 2)
End Function

Function Get_Pipe_X_Length(ByVal pipe As HMITubePolyline)
    Dim x1, x2
    pipe.index = 1
    x1 = pipe.ActualPointLeft
    pipe.index = 2
    x2 = pipe.ActualPointLeft
    Get_Pipe_X_Length = Abs(x2 - x1)
End Function

Function Get_Pipe_Y_Length(ByVal pipe As HMITubePolyline)
    Dim y1, y2
    pipe.index = 1
    y1 = pipe.ActualPointTop
    pipe.index = 2
    y2 = pipe.ActualPointTop
    Get_Pipe_Y_Length = Abs(y2 - y1)
End Function

Function Get_Pipe_X_Min(ByVal pipe As HMITubePolyline)
    Dim i, minX
    
    pipe.index = 1
    minX = pipe.ActualPointLeft
    
    For i = 0 To pipe.PointCount - 1
        pipe.index = i + 1
        If pipe.ActualPointLeft < minX Then
            minX = pipe.ActualPointLeft
        End If 'pipe.ActualPointLeft < minX
    Next i
    
    Get_Pipe_X_Min = minX
End Function

Function Get_Pipe_Y_Min(ByVal pipe As HMITubePolyline)
    Dim i, minY
    
    pipe.index = 1
    minY = pipe.ActualPointTop
    
    For i = 0 To pipe.PointCount - 1
        pipe.index = i + 1
        If pipe.ActualPointTop < minY Then
            minY = pipe.ActualPointTop
        End If 'pipe.ActualPointLeft < minX
    Next i
    
    Get_Pipe_Y_Min = minY
End Function

Sub Remove_Vertex(ByVal pipe As HMITubePolyline, ByVal index As Integer)
    Dim i, j
    Dim coords() As Integer
    Dim numVertices
    
    numVertices = pipe.PointCount
    coords = Get_Coords(pipe)
    
    j = 1
    
    For i = 1 To numVertices
        If i <> index Then
            pipe.index = j
            pipe.ActualPointLeft = coords(0, i - 1)
            pipe.ActualPointTop = coords(1, i - 1)
            j = j + 1
        End If
    Next i
    
    pipe.PointCount = numVertices - 1
    
    If index = 1 Then
        pipe.Left = coords(0, 1)
        pipe.Top = coords(1, 1)
    End If
    
End Sub

Function Get_Closest_Vertices(ByVal pipe1 As HMITubePolyline, ByVal pipe2 As HMITubePolyline)
    Dim vertices(1) As Integer
    Dim i, j, distMin As Integer
    Dim dist11 As Integer, dist12 As Integer, dist21 As Integer, dist22 As Integer
    

    dist11 = Vertex_Distance(pipe1, 1, pipe2, 1)
    
    dist12 = Vertex_Distance(pipe1, 1, pipe2, pipe2.PointCount)

    dist21 = Vertex_Distance(pipe1, pipe1.PointCount, pipe2, 1)
    
    dist22 = Vertex_Distance(pipe1, pipe1.PointCount, pipe2, pipe2.PointCount)

    distMin = dist11
    
    If dist12 < distMin Then
        distMin = dist12
    End If
    
    If dist21 < distMin Then
        distMin = dist21
    End If
    
    If dist22 < distMin Then
        distMin = dist22
    End If
    
    
    If distMin = dist11 Then
        vertices(0) = 1
        vertices(1) = 1
    End If
    
    If distMin = dist12 Then
        vertices(0) = 1
        vertices(1) = pipe2.PointCount
    End If
    
    If distMin = dist21 Then
        vertices(0) = pipe1.PointCount
        vertices(1) = 1
    End If
    
    If distMin = dist22 Then
        vertices(0) = pipe1.PointCount
        vertices(1) = pipe2.PointCount
    End If
    
    Get_Closest_Vertices = vertices
    
End Function

Function Get_Closest_Pipes(ByRef pipeList() As HMITubePolyline)
    Dim pipeListSize As Integer
    Dim closestPipes(1) As Integer
    Dim i, j, pipeMin As Integer, overalMin As Integer
    Dim pipeMinVerts() As Integer
    
    pipeListSize = (UBound(pipeList) - LBound(pipeList)) + 1
    
    pipeMinVerts = Get_Closest_Vertices(pipeList(0), pipeList(1))
    overallMin = Vertex_Distance(pipeList(0), pipeMinVerts(0), pipeList(1), pipeMinVerts(1))
    closestPipes(0) = 0
    closestPipes(1) = 1
    
    For i = 0 To pipeListSize - 2

        For j = i + 1 To pipeListSize - 1
            pipeMinVerts = Get_Closest_Vertices(pipeList(i), pipeList(i + 1))
            pipeMin = Vertex_Distance(pipeList(i), pipeMinVerts(0), pipeList(i + 1), pipeMinVerts(1))

            If pipeMin < overallMin Then
                overallMin = pipeMin
                closestPipes(0) = i
                closestPipes(1) = j
            End If
            
        Next j
    Next i
    
    Get_Closest_Pipes = closestPipes
    
End Function

Function Vertex_Distance(ByVal pipe1 As HMITubePolyline, ByVal index1 As Integer, ByVal pipe2 As HMITubePolyline, ByVal index2 As Integer)
    Dim x1, x2, y1, y2
    Dim distX As Integer, distY As Integer
    
    pipe1.index = index1
    x1 = pipe1.ActualPointLeft
    y1 = pipe1.ActualPointTop
    pipe2.index = index2
    x2 = pipe2.ActualPointLeft
    y2 = pipe2.ActualPointTop
    
    distX = Abs(x1 - x2)
    distY = Abs(y1 - y2)
    Vertex_Distance = Sqr((distX ^ 2) + (distY ^ 2))
End Function

Function Pipe_Total_Length(ByVal pipe As HMITubePolyline)
    Dim i, length
    
    For i = 1 To pipe.PointCount - 1
        length = length + Vertex_Distance(pipe, i, pipe, i + 1)
    Next i
    
    Pipe_Total_Length = length
End Function

Function Pipe_Intersect(pipe1 As HMITubePolyline, point1 As Integer, pipe2 As HMITubePolyline, point2 As Integer)
    Dim p1_x1 As Double, p1_x2 As Double, p1_y1 As Double, p1_y2 As Double
    Dim p2_x1 As Double, p2_x2 As Double, p2_y1 As Double, p2_y2 As Double
    Dim m1 As Double, m2 As Double, b1 As Double, b2 As Double
    Dim p1Vert, p2Vert
    Dim line1, line2
    Dim coord(1) As Integer
    Dim closestVerts() As Integer
    
    If point1 = 1 Then
        line1 = 1
    Else
        line1 = pipe1.PointCount - 1
    End If
    
    If point2 = 1 Then
        line2 = 1
    Else
        line2 = pipe2.PointCount - 1
    End If
    
    pipe1.index = line1
    p1_x1 = pipe1.ActualPointLeft
    p1_y1 = pipe1.ActualPointTop
    pipe1.index = line1 + 1
    p1_x2 = pipe1.ActualPointLeft
    p1_y2 = pipe1.ActualPointTop
    
    pipe2.index = line2
    p2_x1 = pipe2.ActualPointLeft
    p2_y1 = pipe2.ActualPointTop
    pipe2.index = line2 + 1
    p2_x2 = pipe2.ActualPointLeft
    p2_y2 = pipe2.ActualPointTop
    
    If p1_x1 <> p1_x2 Then
        m1 = (p1_y1 - p1_y2) / (p1_x1 - p1_x2)
        b1 = p1_y1 - (m1 * p1_x1)
        p1Vert = False
    Else
        p1Vert = True
    End If
    
    If p2_x1 <> p2_x2 Then
        m2 = (p2_y1 - p2_y2) / (p2_x1 - p2_x2)
        b2 = p2_y1 - (m2 * p2_x1)
        p2Vert = False
    Else
        p2Vert = True
    End If
    
    If (p1Vert And p2Vert) Or ((m1 = m2) And (Not (p1Vert) And Not (p2Vert))) Then
        closestVerts = Get_Closest_Vertices(pipe1, pipe2)
        pipe1.index = closestVerts(0)
        pipe2.index = closestVerts(1)
        coord(0) = pipe1.ActualPointLeft + ((pipe2.ActualPointLeft - pipe1.ActualPointLeft) / 2)
        coord(1) = pipe1.ActualPointTop + ((pipe2.ActualPointTop - pipe1.ActualPointTop) / 2)
    Else
        If p1Vert Or p2Vert Then
            If p1Vert Then
                coord(0) = p1_x1
                coord(1) = (m2 * coord(0)) + b2
            End If
            If p2Vert Then
                coord(0) = p2_x1
                coord(1) = (m1 * coord(0)) + b1
            End If
        Else
            coord(0) = (b2 - b1) / (m1 - m2)
            coord(1) = (m1 * coord(0)) + b1
        End If
    
    End If
    
    Pipe_Intersect = coord
End Function

Sub Pipe_Parallel_Allign(ByVal pipe1 As HMITubePolyline, ByVal point1 As Integer, ByVal pipe2 As HMITubePolyline, ByVal point2 As Integer)
    Dim p1_x1 As Double, p1_x2 As Double, p1_y1 As Double, p1_y2 As Double
    Dim p2_x1 As Double, p2_x2 As Double, p2_y1 As Double, p2_y2 As Double
    Dim m1 As Double, m2 As Double, b1 As Double, b2 As Double
    Dim line1, line2
    Dim bAdjust As Double
    
    If point1 = 1 Then
        line1 = 1
    Else
        line1 = pipe1.PointCount - 1
    End If
    
    If point2 = 1 Then
        line2 = 1
    Else
        line2 = pipe2.PointCount - 1
    End If
    
    pipe1.index = line1
    p1_x1 = pipe1.ActualPointLeft
    p1_y1 = pipe1.ActualPointTop
    pipe1.index = line1 + 1
    p1_x2 = pipe1.ActualPointLeft
    p1_y2 = pipe1.ActualPointTop
    
    pipe2.index = line2
    p2_x1 = pipe2.ActualPointLeft
    p2_y1 = pipe2.ActualPointTop
    pipe2.index = line2 + 1
    p2_x2 = pipe2.ActualPointLeft
    p2_y2 = pipe2.ActualPointTop
    
    If (p1_x1 = p1_x2) And (p2_x1 = p2_x2) And (p1_x1 <> p2_x1) Then
        If Pipe_Total_Length(pipe2) > Pipe_Total_Length(pipe1) Then
            pipe1.Left = pipe2.Left
        Else
            pipe2.Left = pipe1.Left
        End If
    ElseIf (p1_y1 = p1_y2) And (p2_y1 = p2_y2) And (p1_y1 <> p2_y1) Then
        If Pipe_Total_Length(pipe2) > Pipe_Total_Length(pipe1) Then
            pipe1.Top = pipe2.Top
        Else
            pipe2.Top = pipe1.Top
        End If
    Else
        If p1_x1 <> p1_x2 Then
            m1 = (p1_y1 - p1_y2) / (p1_x1 - p1_x2)
            m1 = Round(m1, 5)
            b1 = p1_y1 - (m1 * p1_x1)
        End If
        If p2_x1 <> p2_x2 Then
            m2 = (p2_y1 - p2_y2) / (p2_x1 - p2_x2)
            m2 = Round(m2, 5)
            b2 = p2_y1 - (m2 * p2_x1)
        End If
        
        If (m1 = m2) And (b1 <> b2) And (Not ((p1_x1 <> p1_x2) Or (p2_x1 <> p2))) Then
            If Pipe_Total_Length(pipe2) > Pipe_Total_Length(pipe1) Then
                bAdjust = (b2 - b1) / 2
                
                If m1 <> 0 Then
                    pipe1.Left = pipe1.Left - ((1 / m1) * bAdjust)
                Else
                    pipe1.Left = pipe2.Left
                End If
                
                pipe1.Top = pipe1.Top + bAdjust
            Else
                bAdjust = (b1 - b2) / 2
                
                If m2 <> 0 Then
                    pipe2.Left = pipe2.Left - ((1 / m2) * bAdjust)
                Else
                    pipe2.Left = pipe1.Left
                End If
                pipe2.Top = pipe2.Top + bAdjust
            End If
        End If
    End If
    
End Sub

Function Pipe_Connector_Direction(ByVal pipe As HMITubePolyline, ByVal vertex As Integer)
    'Directions
    '1 = Up
    '2 = Right
    '4 = Down
    '8 = Left
    'E.G 3 = Up + Right, 11 = Up + Right + Left, 15 = Up + Right + Down + Left
    'If only 1, 2, 4, or 8 no connector will be made
    
    Dim baseCoord(1) As Integer
    Dim compCoord(1) As Integer
    Dim direction As Integer
    
    pipe.index = vertex
    baseCoord(0) = pipe.ActualPointLeft
    baseCoord(1) = pipe.ActualPointTop
    
    If vertex > 1 Then
    'Compare vertex before
        pipe.index = vertex - 1
        compCoord(0) = pipe.ActualPointLeft
        compCoord(1) = pipe.ActualPointTop
        direction = direction + Coord_Relative_Direction(baseCoord, compCoord)
    End If
    
    If vertex < pipe.PointCount Then
    'Compare vertex after
        pipe.index = vertex + 1
        compCoord(0) = pipe.ActualPointLeft
        compCoord(1) = pipe.ActualPointTop
        direction = direction + Coord_Relative_Direction(baseCoord, compCoord)
    End If
    
    Pipe_Connector_Direction = direction
    
End Function

Function Coord_Relative_Direction(ByRef baseCoord() As Integer, ByRef compCoord() As Integer)
    'Directions
    '1 = Up
    '2 = Right
    '4 = Down
    '8 = Left
    
    If (baseCoord(0) = compCoord(0)) And Not (baseCoord(1) = compCoord(1)) Then
    'On y axis
    
        If compCoord(1) < baseCoord(1) Then
        'comp up from base
            Coord_Relative_Direction = 1
        ElseIf compCoord(1) > baseCoord(1) Then
        'comp down from base
            Coord_Relative_Direction = 4
        End If
    
    ElseIf (baseCoord(1) = compCoord(1)) And Not (baseCoord(0) = compCoord(0)) Then
    'On x axis
    
        If compCoord(0) < baseCoord(0) Then
        'comp left from base
            Coord_Relative_Direction = 8
        ElseIf compCoord(0) > baseCoord(0) Then
        'comp right from base
            Coord_Relative_Direction = 2
        End If
    
    Else
    'Same point or not vertical or horizontal
        
        Coord_Relative_Direction = 0
        
    End If
End Function

