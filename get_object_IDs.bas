Attribute VB_Name = "Module1"
Option Explicit
Sub OpenRequestDialog()
Attribute OpenRequestDialog.VB_ProcData.VB_Invoke_Func = "Q\n14"

UserForm1.Show vbModeless
 
End Sub

Sub GetObjectNumbers()
Attribute GetObjectNumbers.VB_ProcData.VB_Invoke_Func = " \n14"

'The Robot.Project.Structure Data server manages number of individual servers that correspond to each object type
'(such as the node server, bar server, etc). Before information can be obtained about the model, the script must:
'   1. Obtain a RobotSelection class instance which containes a selection of objects with specific type such as nodes
'    (initial user selection can be of bars, nodes etc; so specific object type must be selected to use in step no 2,_
'    otherwise requesting information of nodes and bars from a node server will throw an error)
'   2. Obtain the specific server class instance which manages the RobotSelection type from step 1
'   3. Use the server method to get a collection of objects
'   4.loop through each object and use parameters/methods as required
'
'the above process needs to be done for each object type seperately, each time creating a single type selection, requesting server object,then getting the objects so;
'the user is first asked to specify which objects to request,the sctipt then check wchich requests were made, makes the neccesary server objects_
' and loops through each request using the above process.


Dim robapp As IRobotApplication
Dim sel As IRobotSelection
Dim obj_server As RobotNodeServer
Dim obj_col As IRobotCollection
Dim oTypeKey As Variant 'IRobotObjectType (key for node,bar,panel)

Dim dUserRequest As Dictionary
Dim cReqServer As Dictionary
Dim server As Variant

Dim ObjNumber() As String
Dim ArrSize As Long
Dim NewArrSize As Long
Dim i As Long
Dim j As Long

'get Robot Application
Set robapp = New RobotApplication

'get user request
Set dUserRequest = get_User_Request()

'get requested servers
Set cReqServer = get_Requested_Server(robapp, dUserRequest)

'initialise array
ReDim ObjNumber(0) As String

'for each requested object type (as dictionary key representing type of object with IRobotObjectType)
For Each oTypeKey In cReqServer.Keys

    'get a selection object of things selected in robot (nodes)
    Set sel = robapp.Project.Structure.Selections.Get(oTypeKey)
    
    'get objects defined by selection
    'collectionRequestedServer.Item(object type) represents the requested object server such as_
    'RobotNodeServer from OMRobot.RobotApplication.Project.Structure.Nodes ### See get_Requested_Server
    Set obj_col = cReqServer.Item(oTypeKey).GetMany(sel)
    
    'do work on nodes
    ArrSize = GetLength(ObjNumber)
    NewArrSize = ArrSize + obj_col.Count - 1
    ReDim Preserve ObjNumber(NewArrSize) As String
    j = 0
    If obj_col.Count <> 0 Then
        For i = ArrSize To NewArrSize
            j = j + 1
            ObjNumber(i) = obj_col.Get(j).Number
        Next i
    End If
    
Next oTypeKey

ActiveCell.Value = Join(ObjNumber, " ")

'End
Set obj_server = Nothing
Set robapp = Nothing
 
End Sub

Function get_User_Request() As Dictionary

Set get_User_Request = New Dictionary

'set IRobotObjectType to true if requested by user

If UserForm1.Nodes_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_NODE, True
If UserForm1.Bars_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_BAR, True
If UserForm1.Panels_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_PANEL, True

End Function

Function get_Requested_Server(oRobApp As IRobotApplication, dUserRequest As Dictionary) As Dictionary

Set get_Requested_Server = New Dictionary
Dim Key As Variant

'for each requested IRobotObjectType thats true, create the correct server object, set it to the one from open revit model and
'add to dictionary returned by fucntion

For Each Key In dUserRequest
    If dUserRequest(Key) = True Then
        Select Case Key
            Case I_OT_NODE
                Dim obj_node_server As RobotNodeServer
                Set obj_node_server = oRobApp.Project.Structure.Nodes
                get_Requested_Server.Add Key, obj_node_server
            Case I_OT_BAR
                Dim obj_bar_server As RobotBarServer
                Set obj_bar_server = oRobApp.Project.Structure.Bars
                get_Requested_Server.Add Key, obj_bar_server
            Case I_OT_PANEL
                Dim obj_panel_server As RobotObjObjectServer
                Set obj_panel_server = oRobApp.Project.Structure.Objects
                get_Requested_Server.Add Key, obj_panel_server
        End Select
    End If
Next Key

End Function

