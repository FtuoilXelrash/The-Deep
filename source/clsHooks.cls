VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents achooks As Decal.achooks
Attribute achooks.VB_VarHelpID = -1


Private Sub AChooks_ChatTextIntercept(ByVal bstrText As String, _
ByVal lColor As Long, ByVal lTarget As Long, bEat As Boolean)

On Error GoTo err_dead


    If lColor = 0 Then
    
    If InStr(bstrText, "All of the nearby disciples begin to chant wildly.") And _
    control.chkFilterDeepSpam.Checked = True Then
    bEat = True
    End If
        
    If (InStr(1, bstrText, "of The Deep", vbTextCompare) <> 0) And _
    control.chkFilterDeepSpam.Checked = True Then
    bEat = True
    End If
       
    If (InStr(1, bstrText, "You give The Deep", vbTextCompare) <> 0) And _
    control.chkFilterDeepSpam.Checked = True Then
    bEat = True
    End If
                
    If (InStr(1, bstrText, "You drop the Gold", vbTextCompare) <> 0) And _
    control.chkFilterDeepSpam.Checked = True Then
    bEat = True
    End If
    End If

    Exit Sub
    
err_dead:
    Hub.SendText "err_AChooks_ChatTextIntercept: " & Err.Description, 1
End Sub




