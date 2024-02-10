VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorldFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents wfilter As DecalFilters.world
Attribute wfilter.VB_VarHelpID = -1

Public Function GetUnStackedItemCount(sItem As String) As Long
On Error GoTo err_dead

Dim wFilIter As DecalFilters.WorldIterator
Dim tmpItem As DecalFilters.IWorldObject

   GetUnStackedItemCount = 0
   If Len(sItem) > 0 Then
      Set wFilIter = wfilter.Inventory 'World object declared WithEvents
      wFilIter.ByName (sItem)
      For Each tmpItem In wFilIter
         GetUnStackedItemCount = GetUnStackedItemCount + 1
      Next
   End If
   
   Set wFilIter = Nothing
   
    Exit Function

err_dead:
    Hub.SendText "GetUnStackedItemCount: " & Err.Description, 1
End Function

Public Function FindGoldsGuid(TheItem As String) As Long
On Error GoTo err_dead

Dim wFilIter As DecalFilters.WorldIterator
Dim tmpItem As DecalFilters.IWorldObject

   If Len(TheItem) > 0 Then
      Set wFilIter = wfilter.Inventory 'World object declared WithEvents
      wFilIter.ByName (TheItem)
      For Each tmpItem In wFilIter
        FindGoldsGuid = tmpItem.Guid
'        WriteToChat "(DEBUG) - The Item Guid: " & FindGoldsGuid, 10
        Next

   End If
   Set wFilIter = Nothing
    Exit Function

err_dead:
    Hub.SendText "FindGoldsGuid: " & Err.Description, 1
End Function

Public Sub CountGoldPeices()
On Error GoTo err_dead

Dim CountTotalGolds As Integer

Dim GoldFin As Integer
Dim GoldFinXP As Double

Dim GoldTent As Integer
Dim GoldTentXP As Double

Dim GoldTooth As Integer
Dim GoldToothXP As Double

Dim GoldEgg As Integer
Dim GoldEggXP As Double

Dim TotalGoldXP As Double

GoldFin = GetUnStackedItemCount("Gold Shallows Shredder Fin")
GoldFinXP = (GoldFin * 6500000)

GoldTent = GetUnStackedItemCount("Gold Niffis Tentacle")
GoldTentXP = (GoldTent * 8700000)

GoldTooth = GetUnStackedItemCount("Gold Moarsman Tooth")
GoldToothXP = (GoldTooth * 6500000)

GoldEgg = GetUnStackedItemCount("Gold Remoran Eggs")
GoldEggXP = (GoldEgg * 11400000)

CountTotalGolds = GoldFin + GoldTent + GoldTooth + GoldEgg
TotalGoldXP = GoldFinXP + GoldToothXP + GoldTentXP + GoldEggXP

Call control.ChangeStaticText(view.MainView.control("lblGoldFin"), GoldFin)
Call control.ChangeStaticText(view.MainView.control("lblGoldFinXP"), GoldFinXP)

Call control.ChangeStaticText(view.MainView.control("lblGoldTentacle"), GoldTent)
Call control.ChangeStaticText(view.MainView.control("lblGoldTentacleXP"), GoldTentXP)

Call control.ChangeStaticText(view.MainView.control("lblGoldTooth"), GoldTooth)
Call control.ChangeStaticText(view.MainView.control("lblGoldToothXP"), GoldToothXP)

Call control.ChangeStaticText(view.MainView.control("lblGoldEgg"), GoldEgg)
Call control.ChangeStaticText(view.MainView.control("lblGoldEggXP"), GoldEggXP)

Call control.ChangeStaticText(view.MainView.control("lblDeepToken"), GetUnStackedItemCount("Servant of The Deep Token"))

'WriteToChat "CountTotalGolds: " & CountTotalGolds, 10
'WriteToChat "CountTotalGolds: " & view.MainView.control("lblGoldFin").ChildCount, 10
Call control.ChangeStaticText(view.MainView.control("lblTotalGoldsCount"), CountTotalGolds)
Call control.ChangeStaticText(view.MainView.control("lblTotalXP"), TotalGoldXP)


    Exit Sub
    
err_dead:
    Hub.SendText "CountGoldPeices: " & Err.Description, 1
End Sub

Public Function NameFromGuid(ByVal ObjectGUID As Long) As String
On Error GoTo err_dead

    Dim wObj As WorldObject
    Set wObj = wfilter(ObjectGUID)

    NameFromGuid = wObj.Strings(keyName)

    Exit Function
err_dead:
    Hub.SendText "NameFromGuid: " & Err.Description, 1
End Function
