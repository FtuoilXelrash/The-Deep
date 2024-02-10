VERSION 5.00
Begin VB.Form TimerForm 
   Caption         =   "TimerForm"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "TimerForm"
   ScaleHeight     =   1155
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer DeepTurninTimer 
      Enabled         =   0   'False
      Interval        =   190
      Left            =   300
      Top             =   120
   End
End
Attribute VB_Name = "TimerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DeepTurninTimer_Timer()

'    WriteToChat "(DEBUG) - *DEEP TIMER FIRED*", 3
    
    Dim PeiceGuidNum As Long
    Dim PeiceCount As Integer
    
    If TheDeepTrophy = "" And _
    control.chkGoldFin.Checked = True Then
    TheDeepTrophy = "Gold Shallows Shredder Fin"
    End If
    
    If TheDeepTrophy = "" And _
    control.chkGoldTentacle.Checked = True Then
    TheDeepTrophy = "Gold Niffis Tentacle"
    End If
    
    If TheDeepTrophy = "" And _
    control.chkGoldTooth.Checked = True Then
    TheDeepTrophy = "Gold Moarsman Tooth"
    End If
    
    If TheDeepTrophy = "" And _
    control.chkGoldEgg.Checked = True Then
    TheDeepTrophy = "Gold Remoran Eggs"
    End If
    
    If TheDeepTrophy = "" And _
    control.chkDeepToken.Checked = True Then
    TheDeepTrophy = "Servant of The Deep Token"
    End If
    
    PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
    
'   CHECK TO MAKE SURE THEY DIDNT HIT STOP!
If DeepTurnIn_OK = True Then
    
'   Gold Shallows Shredder Fin
    If TheDeepTrophy = "Gold Shallows Shredder Fin" Then
    
        If control.chkGoldFin.Checked = True Then

            If PeiceCount > 0 Then
            PeiceGuidNum = world.FindGoldsGuid(TheDeepTrophy)
            Call hook.achooks.GiveItem(PeiceGuidNum, DeepTurninTarget)
            world.CountGoldPeices
            
            Else
            world.CountGoldPeices
            WriteToChat TheDeepTrophy & ": Completed", 10
            
            TheDeepTrophy = "Gold Niffis Tentacle"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
            Else
            
            TheDeepTrophy = "Gold Niffis Tentacle"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
        End If
    End If
    
'   Gold Niffis Tentacle
    If TheDeepTrophy = "Gold Niffis Tentacle" Then
    
        If control.chkGoldTentacle.Checked = True Then
        
            If PeiceCount > 0 Then
            PeiceGuidNum = world.FindGoldsGuid(TheDeepTrophy)
            Call hook.achooks.GiveItem(PeiceGuidNum, DeepTurninTarget)
            world.CountGoldPeices
            Else
            world.CountGoldPeices
            WriteToChat TheDeepTrophy & ": Completed", 10
            
            TheDeepTrophy = "Gold Moarsman Tooth"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
            Else
            TheDeepTrophy = "Gold Moarsman Tooth"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
        
        End If
        
'   Gold Moarsman Tooth
    If TheDeepTrophy = "Gold Moarsman Tooth" Then
    
        If control.chkGoldTooth.Checked = True Then
        
            If PeiceCount > 0 Then
            PeiceGuidNum = world.FindGoldsGuid(TheDeepTrophy)
            Call hook.achooks.GiveItem(PeiceGuidNum, DeepTurninTarget)
            world.CountGoldPeices
            Else
            world.CountGoldPeices
            WriteToChat TheDeepTrophy & ": Completed", 10
            
            TheDeepTrophy = "Gold Remoran Eggs"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
            Else
            TheDeepTrophy = "Gold Remoran Eggs"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
        
        End If
        
'   Gold Remoran Eggs
    If TheDeepTrophy = "Gold Remoran Eggs" Then
    
        If control.chkGoldTooth.Checked = True Then
        
            If PeiceCount > 0 Then
            PeiceGuidNum = world.FindGoldsGuid(TheDeepTrophy)
            Call hook.achooks.GiveItem(PeiceGuidNum, DeepTurninTarget)
            world.CountGoldPeices
            Else
            world.CountGoldPeices
            WriteToChat TheDeepTrophy & ": Completed", 10
            
            TheDeepTrophy = "Servant of The Deep Token"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
            Else
            TheDeepTrophy = "Servant of The Deep Token"
            PeiceCount = world.GetUnStackedItemCount(TheDeepTrophy)
            End If
        
        End If
        
'   Servant of the Deep Token
    If TheDeepTrophy = "Servant of The Deep Token" Then
    
        If control.chkGoldTooth.Checked = True Then
        
            If PeiceCount > 0 Then
            PeiceGuidNum = world.FindGoldsGuid(TheDeepTrophy)
            Call hook.achooks.GiveItem(PeiceGuidNum, DeepTurninTarget)
            world.CountGoldPeices
            Else
            world.CountGoldPeices
            WriteToChat TheDeepTrophy & ": Completed", 10
            DeepTurnIn_OK = False
            End If
            Else
            DeepTurnIn_OK = False
            End If
        
        End If
        Else
    world.CountGoldPeices
    control.btnStartDeepTrade.Text = "START"
    control.btnStartDeepTrade.TextColor = "65280"
    TimerForm.DeepTurninTimer.Enabled = False
End If

End Sub

Public Sub ExitGoldTurnin()
    TimerForm.DeepTurninTimer.Enabled = False
    world.CountGoldPeices
    control.btnStartDeepTrade.Text = "START"
    control.btnStartDeepTrade.TextColor = "65280"
End Sub

