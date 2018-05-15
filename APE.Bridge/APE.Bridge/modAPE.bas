Attribute VB_Name = "modAPE"
Option Explicit

Private Const strProgIDBridgeToAPE = "APE.Bridge.Ax"
Private BridgeToAPE As Object

Private Sub CreateBridgeToAPE()
    On Error GoTo CreateError
    Set BridgeToAPE = CreateObject(strProgIDBridgeToAPE)
    On Error GoTo 0
Exit Sub

CreateError:
    Resume RaiseError
RaiseError:
    'Raise a more friendly error than the default
    Call Err.Raise(429, Description:="Failed to create " + strProgIDBridgeToAPE + " object")
End Sub

Private Function GetLabelContainerHandle(containerHandle As Long, ctrl As Control)
    On Error GoTo HandleError
    GetLabelContainerHandle = ctrl.Container.hWnd
Exit Function

HandleError:
    GetLabelContainerHandle = containerHandle
End Function

Public Sub AddObjectToBridgeToAPE(containerHandle As Long, containerName As String, containerControl As Object, containerControlCollection As Object)
    Dim ctrl As Control
    On Error Resume Next
    
    If BridgeToAPE Is Nothing Then
        Call CreateBridgeToAPE
    End If
    
    Call BridgeToAPE.AddItem(ObjPtr(containerControl), containerHandle, containerHandle, containerName, TypeName(containerControl), containerControl, False)
    
    For Each ctrl In containerControlCollection
        If TypeName(ctrl) = "Label" Then
            Call BridgeToAPE.AddItem(ObjPtr(ctrl), containerHandle, GetLabelContainerHandle(containerHandle, ctrl), ctrl.Name, TypeName(ctrl), ctrl, True)
        Else
            Call BridgeToAPE.AddItem(ObjPtr(ctrl), containerHandle, ctrl.hWnd, ctrl.Name, TypeName(ctrl), ctrl, False)
        End If
    Next ctrl
End Sub

Public Sub RemoveContainerFromBridgeToAPE(containerHandle As Long)
    If BridgeToAPE Is Nothing Then
        Exit Sub
    End If
    
    Call BridgeToAPE.RemoveAllItemsFromContainer(containerHandle)
End Sub
