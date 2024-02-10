Attribute VB_Name = "Vars"
Option Explicit

Public boolSettingsLoaded       As Boolean
Public boolLoggedIn             As Boolean
Public bLoginCompleted          As Boolean
Global DEBUG_MODE               As Boolean

Public PluginSite               As Decal.IPluginSite2
Public pluginSiteOld            As DecalPlugins.IPluginSite
Public view                     As clsViewHandler
Public control                  As clsControlHandler
Public hook                     As New clsHooks
Public Csf                      As New clsCharStats
Public world                    As New clsWorldFilter
Public echo                     As New clsNetEcho

'   REGISTRY
Public Saved                    As String
Public reg                      As RegOp

'   THE DEEP
Public DeepTurnIn_OK            As Boolean
Public DeepTurninTarget         As Long
Public TheDeepTrophy            As String
Public TurnIntoGUID             As Long

'   STUFF
Public PlayerGUID               As Long

