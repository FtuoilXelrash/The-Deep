VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'   ABOUT
Private WithEvents lblAboutBuildDate As DecalControls.StaticText
Attribute lblAboutBuildDate.VB_VarHelpID = -1
Public imgLogo As DecalControls.FixedLayout
Public imgBGPic As DecalControls.FixedLayout
Public WithEvents chkGoldFin As DecalControls.CheckBox
Attribute chkGoldFin.VB_VarHelpID = -1
Public WithEvents lblGoldFin As DecalControls.StaticText
Attribute lblGoldFin.VB_VarHelpID = -1
Public WithEvents chkGoldTentacle As DecalControls.CheckBox
Attribute chkGoldTentacle.VB_VarHelpID = -1
Public WithEvents chkGoldTooth As DecalControls.CheckBox
Attribute chkGoldTooth.VB_VarHelpID = -1
Public WithEvents chkGoldEgg As DecalControls.CheckBox
Attribute chkGoldEgg.VB_VarHelpID = -1
Public WithEvents chkDeepToken As DecalControls.CheckBox
Attribute chkDeepToken.VB_VarHelpID = -1
Public WithEvents btnRefreshGolds As DecalControls.PushButton
Attribute btnRefreshGolds.VB_VarHelpID = -1
Public WithEvents btnStartDeepTrade As DecalControls.PushButton
Attribute btnStartDeepTrade.VB_VarHelpID = -1
Public WithEvents btnSelectATarget As DecalControls.PushButton
Attribute btnSelectATarget.VB_VarHelpID = -1
Public WithEvents lblCurrentTargetTxt As DecalControls.StaticText
Attribute lblCurrentTargetTxt.VB_VarHelpID = -1

Public WithEvents chkFilterDeepSpam As DecalControls.CheckBox
Attribute chkFilterDeepSpam.VB_VarHelpID = -1

' Initialize the controls used in the Plug-in GUI
Public Function InitControls() As Boolean
On Error GoTo err_dead
    With view.MainView
       
    Set lblAboutBuildDate = .control("lblAboutBuildDate")
    Set imgLogo = .control("imgLogo")
    Set imgBGPic = .control("imgBGPic")
    Set chkGoldFin = .control("chkGoldFin")
    Set lblGoldFin = .control("lblGoldFin")
    Set chkGoldTentacle = .control("chkGoldTentacle")
    Set chkGoldTooth = .control("chkGoldTooth")
    Set chkGoldEgg = .control("chkGoldEgg")
    Set chkDeepToken = .control("chkDeepToken")
    Set btnRefreshGolds = .control("btnRefreshGolds")
    Set btnStartDeepTrade = .control("btnStartDeepTrade")
    Set btnSelectATarget = .control("btnSelectATarget")
    Set lblCurrentTargetTxt = .control("lblCurrentTargetTxt")
    Set chkFilterDeepSpam = .control("chkFilterDeepSpam")

    End With
    Call SetRandomLogo
    Call LoadSettings
    
    Exit Function
    
err_dead:
    Hub.SendText "err_InitControls: " & Err.Description, 1
End Function

Public Sub SetRandomLogo()
       Dim TheImage As IImageCache
'&H6001157
'&H600118A
'       Call hub.pluginSiteOld.LoadBitmapPortal((&H6001157), TheImage)
'       Call hub.pluginSiteOld.LoadBitmapPortal(Choose(Int(4 * Rnd + 1), &H600116E, &H6001162, &H600117A, &H6001163), TheImage)
'       TheImage = hub.pluginSiteOld.LoadBitmapPortal(&H600118A)
       
       Set imgLogo.Background = pluginSiteOld.LoadBitmapPortal(&H600118A)
       Set imgBGPic.Background = pluginSiteOld.LoadBitmapPortal(&H6006067)
       Set TheImage = Nothing
       Exit Sub
End Sub

Public Function LoadSettings()
On Error GoTo err_dead

    Call LoadRegSetting
    Call ChangeStaticText(view.MainView.control("lblAboutBuildDate"), "" & _
    "Plug-in Version: " & Hub.AppVersion)
    Exit Function
    
err_dead:
    Hub.SendText "err_LoadSettings: " & Err.Description, 1
End Function

Public Sub LoadRegSetting()
On Error GoTo err_dead

'  Set reg = New RegOp
  With reg
    .Root = HKEY_LOCAL_MACHINE
    .Key = "Software\Decal Plugins\DEEP"
    control.chkGoldFin.Checked = CBool(.Value("chkGoldFin"))
    control.chkGoldTentacle.Checked = CBool(.Value("chkGoldTentacle"))
    control.chkGoldTooth.Checked = CBool(.Value("chkGoldTooth"))
    control.chkGoldEgg.Checked = CBool(.Value("chkGoldEgg"))
    control.chkDeepToken.Checked = CBool(.Value("chkDeepToken"))
    control.chkFilterDeepSpam.Checked = CBool(.Value("chkFilterDeepSpam"))
    
      End With
      Exit Sub
err_dead:
    Hub.SendText "LoadRegSetting: " & Err.Description, 1
End Sub

Public Sub ChangeStaticText(ByRef TheControl As DecalControls.StaticText, ByVal TheText As String)
    TheControl.Text = TheText
End Sub

Private Sub chkGoldFin_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkGoldFin") = CInt(bChecked)
  
    If chkGoldFin.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Gold Fin Turnover", Saved)
End Sub

Private Sub chkGoldTentacle_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkGoldTentacle") = CInt(bChecked)
  
    If chkGoldTentacle.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Gold Tentacle Turnover", Saved)
End Sub

Private Sub chkGoldTooth_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkGoldTooth") = CInt(bChecked)
  
    If chkGoldTooth.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Gold Tooth Turnover", Saved)
End Sub

Private Sub chkGoldEgg_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkGoldEgg") = CInt(bChecked)
  
    If chkGoldEgg.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Gold Egg Turnover", Saved)
End Sub

Private Sub chkDeepToken_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkDeepToken") = CInt(bChecked)
  
    If chkDeepToken.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Deep Token Turnover", Saved)
End Sub

'   CHANGE REG SETTINGS AND ALERT USER WITH MESSAGE
Public Sub CheckSaved(ByVal WhatChecked As String, Optional ByVal OnOff As String)
    WriteToChat "The Change to " & WhatChecked & " setting was saved " & OnOff & "", 13
End Sub

Public Sub btnSelectATarget_Accepted(ByVal nID As Long)
'    WriteToChat "(DEBUG) - Select A Target", 10

    If hook.achooks.CurrentSelection = PlayerGUID Then
    WriteToChat "You can not hand stuff to yourself! Please select a vailid target!", 10
    Exit Sub
    End If
    
    If hook.achooks.CurrentSelection = 0 Then
    WriteToChat "You MUST Select a valid target!", 10
    Else
    TurnIntoGUID = hook.achooks.CurrentSelection
    
    Call control.ChangeStaticText(control.lblCurrentTargetTxt, _
    world.NameFromGuid(hook.achooks.CurrentSelection))
    End If
    
End Sub

Public Sub btnStartDeepTrade_Accepted(ByVal nID As Long)
    If DeepTurnIn_OK = False Then
    control.btnStartDeepTrade.Text = "STOP"
    control.btnStartDeepTrade.TextColor = "255"
    control.HandOverGolds (TurnIntoGUID)
    Else
    control.btnStartDeepTrade.Text = "START"
    control.btnStartDeepTrade.TextColor = "65280"
    DeepTurnIn_OK = False
    WriteToChat "Deep Turnin Stopped!", 10
    End If
End Sub

Public Sub btnRefreshGolds_Accepted(ByVal nID As Long)
    world.CountGoldPeices
End Sub

Public Sub HandOverGolds(ByVal GoldsTarget As String)
    
    DeepTurnIn_OK = True
    
    Call world.CountGoldPeices
    WriteToChat "Deep Exchange Started Target: " & world.NameFromGuid(GoldsTarget), 10
    
    If GoldsTarget = 0 Or GoldsTarget = "" Or GoldsTarget = "0" Then
    WriteToChat "YOU MUST SELECT A TARGET TO HAND TO!!!", 10
    control.btnStartDeepTrade.Text = "START"
    control.btnStartDeepTrade.TextColor = "65280"
    DeepTurnIn_OK = False
    Else
    DeepTurninTarget = GoldsTarget
    TheDeepTrophy = ""
    TimerForm.DeepTurninTimer.Enabled = True
    End If

End Sub

Private Sub chkFilterDeepSpam_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkFilterDeepSpam") = CInt(bChecked)
  
    If chkFilterDeepSpam.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Filter Deep Spam", Saved)
End Sub


