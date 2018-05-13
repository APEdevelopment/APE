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
    
    Call BridgeToAPE.AddItem(containerHandle, containerName, TypeName(containerControl), containerControl, False)
    
    For Each ctrl In containerControlCollection
        If TypeName(ctrl) = "Label" Then
            Call BridgeToAPE.AddItem(GetLabelContainerHandle(containerHandle, ctrl), ctrl.Name, TypeName(ctrl), ctrl, True)
        Else
            Call BridgeToAPE.AddItem(ctrl.hWnd, ctrl.Name, TypeName(ctrl), ctrl, False)
        End If
    Next ctrl
End Sub

Public Sub RemoveObjectFromBridgeToAPE(containerHandle As Long, containerName As String, containerControl As Object, containerControlCollection As Object)
    Dim ctrl As Control
    On Error Resume Next
    
    If BridgeToAPE Is Nothing Then
        Exit Sub
    End If
    
    Call BridgeToAPE.RemoveItem(containerHandle, containerName, TypeName(containerControl), False)
    
    For Each ctrl In containerControlCollection
        If TypeName(ctrl) = "Label" Then
            Call BridgeToAPE.RemoveItem(GetLabelContainerHandle(containerHandle, ctrl), ctrl.Name, TypeName(ctrl), True)
        Else
            Call BridgeToAPE.RemoveItem(ctrl.hWnd, ctrl.Name, TypeName(ctrl), False)
        End If
    Next ctrl
End Sub
