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

Public Sub AddObjectToBridgeToAPE(containerHandle As Long, containerName As String, containerControl As Object, containerControlCollection As Object)
On Error Resume Next    'For debuging remove this line its just there to make sure if things go wrong it doesn't crash the application
    Dim ctrl As VB.Control
    
    If BridgeToAPE Is Nothing Then
        Call CreateBridgeToAPE
        If BridgeToAPE Is Nothing Then
            Exit Sub
        End If
    End If
    
    Call BridgeToAPE.AddItem(ObjPtr(containerControl), ObjPtr(containerControl), containerHandle, containerName, containerControl, False)
    
    For Each ctrl In containerControlCollection
        Call AddControlToBridge(containerHandle, containerControl, ctrl)
    Next ctrl
End Sub

Public Sub RemoveContainerFromBridgeToAPE(containerControl As Object)
On Error Resume Next    'For debuging remove this line its just there to make sure if things go wrong it doesn't crash the application
    If BridgeToAPE Is Nothing Then
        Exit Sub
    End If
    
    Call BridgeToAPE.RemoveAllItemsFromContainer(ObjPtr(containerControl))
End Sub

Private Sub CreateBridgeToAPE()
    On Error Resume Next
    Set BridgeToAPE = CreateObject(strProgIDBridgeToAPE)
End Sub

Private Function GetRenderedContainerHandle(containerHandle As Long, containerControl As Object, ctrl As Control)
    If TypeOf ctrl.Container Is VB.Control Then
        GetRenderedContainerHandle = ctrl.Container.hWnd
    Else
        GetRenderedContainerHandle = containerHandle
    End If
End Function

Private Sub AddControlToBridge(containerHandle As Long, containerControl As Object, ctrl As Object)
    Dim extender As VB.VBControlExtender
    
    Select Case True
        'Intrinsic Non-GUI controls (ignore them)
        Case TypeOf ctrl Is VB.App
        Case TypeOf ctrl Is VB.Clipboard
        Case TypeOf ctrl Is VB.Data
        Case TypeOf ctrl Is VB.Global
        Case TypeOf ctrl Is VB.Licenses
        Case TypeOf ctrl Is VB.OLE
        Case TypeOf ctrl Is VB.Printer
        Case TypeOf ctrl Is VB.Screen
        Case TypeOf ctrl Is VB.Timer
        
        'VBControlExtender hosting a custom control
        Case TypeOf ctrl Is VB.VBControlExtender
            Set extender = ctrl
            Call BridgeToAPE.AddItem(ObjPtr(extender.object), ObjPtr(containerControl), 0, extender.Name, extender.object, False)
            
        'Rendered Intrinsic GUI controls
        Case TypeOf ctrl Is VB.Image, _
             TypeOf ctrl Is VB.Label, _
             TypeOf ctrl Is VB.Line, _
             TypeOf ctrl Is VB.Menu, _
             TypeOf ctrl Is VB.Shape
            Call BridgeToAPE.AddItem(ObjPtr(ctrl), ObjPtr(containerControl), GetRenderedContainerHandle(containerHandle, containerControl, ctrl), ctrl.Name, ctrl, True)
            
        'Intrinsic GUI controls
        Case TypeOf ctrl Is VB.CheckBox, _
             TypeOf ctrl Is VB.ComboBox, _
             TypeOf ctrl Is VB.CommandButton, _
             TypeOf ctrl Is VB.ListBox, _
             TypeOf ctrl Is VB.DirListBox, _
             TypeOf ctrl Is VB.DriveListBox, _
             TypeOf ctrl Is VB.FileListBox, _
             TypeOf ctrl Is VB.Form, _
             TypeOf ctrl Is VB.Frame, _
             TypeOf ctrl Is VB.HScrollBar, _
             TypeOf ctrl Is VB.MDIForm, _
             TypeOf ctrl Is VB.OptionButton, _
             TypeOf ctrl Is VB.PictureBox, _
             TypeOf ctrl Is VB.PropertyPage, _
             TypeOf ctrl Is VB.TextBox, _
             TypeOf ctrl Is VB.UserDocument, _
             TypeOf ctrl Is VB.VScrollBar
            Call BridgeToAPE.AddItem(ObjPtr(ctrl), ObjPtr(containerControl), ctrl.hWnd, ctrl.Name, ctrl, False)
            
        'Case TypeOf ctrl Is VB.UserControl
        'Case TypeOf ctrl Is VB.Control
    End Select
End Sub

