VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents CSFilter As DecalFilters.CharacterStats
Attribute CSFilter.VB_VarHelpID = -1

Private Sub CSFilter_LoginComplete()
On Error GoTo err_dead

' Load locations, wand guid
'   SET THE DEBUG FLAG
'    DEBUG_MODE = True
    DEBUG_MODE = False
 
'  SET LOGINCOMPLETED FLAG
    bLoginCompleted = True
 
'   UPDATE THE DEEP TAB
    Call world.CountGoldPeices
     
'   WELCOME KIT
WriteToChat "THE DEEP Ver. " & Hub.AppVersion & " By: Ftuoil Xelrash of Darktide!", 4
WriteToChat "Welcome to the world of " & Csf.CSFilter.Server & "! There is currently " & Csf.CSFilter.ServerPopulation & " players connected at this time.", 1
 
PlayerGUID = Csf.CSFilter.Character

'WriteToChat "(DEBUG) - My GUID is: " & PlayerGUID, 10


    Exit Sub
    
err_dead:
    Hub.SendText "err_IPlugin2_Terminate", 1
End Sub
