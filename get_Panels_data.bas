Attribute VB_Name = "Module1"
Option Explicit
Sub Main()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oRobotApp As New RobotApplication
    Dim oProject As IRobotProject
    Dim dProjectAdata As Dictionary
    
    'Setup
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    oRobotApp.Visible = True
    oRobotApp.Interactive = False
    
    'Get Project A
    Set oProject = oRobotApp.project
    
    'Get Project A Information
    Set dProjectAdata = get_Project_information(oProject)
    'oProject.Close
    
    'print information
    PrintPanelInfo dProjectAdata
    
    'Error Handling
    Debug.Print ("finished")
    oRobotApp.Visible = True
    oRobotApp.Interactive = True

End Sub

Function get_Project_information(proj As RobotProject) As Dictionary
    
    Dim objPanelServer As RobotObjObjectServer
    Dim panelData As Dictionary
    Dim Selection As RobotSelection
    Dim panel_col As RobotObjObjectCollection

    Set objPanelServer = proj.Structure.Objects
    Set get_Project_information = New Dictionary

    Set Selection = proj.Structure.Selections.Create(I_OT_PANEL)
    Selection.FromText "all"
    Set panelData = get_panel_information(objPanelServer.GetMany(Selection))

    get_Project_information.Add "PANELS", panelData

End Function

Function get_panel_information(panel_objects As RobotObjObjectCollection) As Dictionary

    Set get_panel_information = New Dictionary
    Dim vPanel_properties(6) As Variant
    Dim panel As RobotObjObject
    Dim panelPart As IRobotObjPart
    Dim panelNodes() As Variant
    Dim panelNode As Object
    
    Dim i As Long
    Dim j As Long
    
    'loop over all nodes in the collection
    For i = 1 To panel_objects.Count
        
        'get the panel object
        Set panel = panel_objects.get(i)
    
        'get panel thickness
        If panel.HasLabel(I_LT_PANEL_THICKNESS) = True Then
            If panel.GetLabel(I_LT_PANEL_THICKNESS).Name = "ATK_Wall_ConcRC_RC3240_WRC_400_01" Or _
                panel.GetLabel(I_LT_PANEL_THICKNESS).Name = "ATK_Wall_ConcRC_RC3240_WRC_500_01" Then
                    
                    'get panel number
                    vPanel_properties(0) = panel.Number
                    
                    'get panel thickness
                    vPanel_properties(1) = panel.GetLabel(I_LT_PANEL_THICKNESS).Name
                    
                    'get panel point count
                    Set panelPart = panel.GetPart(1)
                    vPanel_properties(2) = panelPart.ModelPoints.Count
                                        
                    'create array of point coordinates
                    ReDim panelNodes(1 To vPanel_properties(2))
                    For j = 1 To UBound(panelNodes)
                        Set panelNode = panel.Main.DefPoints.get(j)
                        panelNodes(j) = Array(panelNode.x, panelNode.Y, panelNode.Z)
                    Next j
                    
                    'add points to array
                    vPanel_properties(3) = panelNodes
                    
                    'add array to dictionary
                    get_panel_information.Add vPanel_properties(0), vPanel_properties
            End If
        End If
        
    'and move onto the next panel
    Next i

End Function

Function PrintPanelInfo(Info As Dictionary)

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim panelID As Variant
    Dim panel As Variant
    
    Dim col As Integer
    
    'Setup
    Set wb = ThisWorkbook
    Set ws = ActiveSheet

    i = 0
    
    For Each panelID In Info("PANELS").Keys
        panel = Info("PANELS")(panelID)
        For j = 0 To 3
            If j = 3 Then
            col = 3
            For k = 1 To UBound(panel(3))
                For m = 0 To 2
                    ws.Range("D4").Offset(col, i).Value = panel(j)(k)(m)
                    col = col + 1
                Next m
            Next k
            Else
                ws.Range("D4").Offset(j, i).Value = panel(j)
            End If
        Next j
        i = i + 1
    Next panelID

End Function

