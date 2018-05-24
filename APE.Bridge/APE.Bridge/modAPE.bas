Attribute VB_Name = "modAPE"
Option Explicit
'
'Copyright 2018 David Beales
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
Private Const strProgIDBridgeToAPE = "APE.Bridge.Ax"
Private BridgeToAPE As Object

Private Sub CreateBridgeToAPE()
    On Error Resume Next
    Set BridgeToAPE = CreateObject(strProgIDBridgeToAPE)
End Sub

Private Function GetLabelContainerHandle(containerHandle As Long, ctrl As Control)
    On Error GoTo HandleError
    GetLabelContainerHandle = ctrl.Container.hWnd
    Exit Function
HandleError:
    GetLabelContainerHandle = containerHandle
End Function

Private Function GetName(ctrl As Object) As String
    On Error GoTo NameFromControlError
    GetName = ctrl.Name
    Exit Function
NameFromControlError:
    GetName = ""
End Function

Private Function GetHandle(ctrl As Object) As Long
    On Error GoTo hWndFromControlError
    GetHandle = ctrl.hWnd
    Exit Function
hWndFromControlError:
    Resume TryWindow
TryWindow:
    On Error GoTo WindowFromControlError
    GetHandle = ctrl.Window
    Exit Function
WindowFromControlError:
    Resume TryHandle
TryHandle:
    On Error GoTo HandleFromControlError
    GetHandle = ctrl.Handle
    Exit Function
HandleFromControlError:
    GetHandle = 0
End Function

Public Sub AddObjectToBridgeToAPE(containerHandle As Long, containerName As String, containerControl As Object, containerControlCollection As Object)
    Dim ctrl As Control
    Dim extender As VBControlExtender
    
    If BridgeToAPE Is Nothing Then
        Call CreateBridgeToAPE
        If BridgeToAPE Is Nothing Then
            Exit Sub
        End If
    End If
    
    'Debug.Print "added: " & containerName & " " & ObjPtr(containerControl)
    Call BridgeToAPE.AddItem(ObjPtr(containerControl), containerHandle, containerHandle, containerName, TypeName(containerControl), containerControl, False)
    
    For Each ctrl In containerControlCollection
        If TypeName(ctrl) = "Label" Then
            'Debug.Print "added label: " & GetName(ctrl) & " " & ObjPtr(ctrl)
            Call BridgeToAPE.AddItem(ObjPtr(ctrl), containerHandle, GetLabelContainerHandle(containerHandle, ctrl), GetName(ctrl), TypeName(ctrl), ctrl, True)
        Else
            If TypeOf ctrl Is VBControlExtender Then
                Set extender = ctrl
                'Debug.Print "added extender: " & extender.Name & " " & ObjPtr(extender.object)
                Call BridgeToAPE.AddItem(ObjPtr(extender.object), containerHandle, GetHandle(extender.object), extender.Name, TypeName(extender.object), extender.object, False)
            Else
                'Debug.Print "added object: " & GetName(ctrl) & " " & ObjPtr(ctrl)
                Call BridgeToAPE.AddItem(ObjPtr(ctrl), containerHandle, GetHandle(ctrl), GetName(ctrl), TypeName(ctrl), ctrl, False)
            End If
        End If
    Next ctrl
End Sub

Public Sub RemoveContainerFromBridgeToAPE(containerHandle As Long)
    If BridgeToAPE Is Nothing Then
        Exit Sub
    End If
    
    Call BridgeToAPE.RemoveAllItemsFromContainer(containerHandle)
End Sub

