VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents MainView As DecalPlugins.view
Attribute MainView.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Dim Schema As String
      Dim i%
        Schema = ViewSchemaOld()
   
   ' Initialize Decal Controls
   Set MainView = pluginSiteOld.LoadView(Schema)
End Sub

Public Function ViewSchemaOld() As String
On Error GoTo err_dead

    ViewSchemaOld = "<?xml version=""1.0""?>"
    ViewSchemaOld = ViewSchemaOld & "<view icon=""8316"" title=""THE DEEP"" top=""55"" left=""3"" width=""215"" height=""280"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Notebook"" name=""THE DEEP"">"
               
    ViewSchemaOld = ViewSchemaOld & "<page label=""OPTIONS"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""DeepOpts"" clipped="""">"
                      
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""imgBGPic"" id=""24679"" left=""0"" top=""5"" width=""215"" height=""220"">"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblDeepCoords"" text=""The Deep 77.8N, 67.1E"" top=""1"" left=""35"" width=""130"" height=""14"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnDevLine"" left=""0"" top=""15"" width=""215"" height=""5"" textcolor="""" text=""""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldFin"" text=""0"" top=""25"" left=""2"" width=""40"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Button"" name=""IcoGoldFin"" left=""30"" top=""25"" width=""16"" height=""16"" icon=""25100"" pressedicon=""25100""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkGoldFin"" left=""55"" top=""25"" width=""20"" height=""20"" text=""""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldFinXP"" text=""0"" top=""25"" left=""75"" width=""150"" height=""16"" fontsize=""16"" textcolor=""16777215"" fontstyle=""bold""/>"

    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldTentacle"" text=""0"" top=""40"" left=""2"" width=""40"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Button"" name=""IcoGoldTentacle"" left=""30"" top=""40"" width=""16"" height=""16"" icon=""25101"" pressedicon=""25101""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkGoldTentacle"" left=""55"" top=""40"" width=""20"" height=""20"" text=""""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldTentacleXP"" text=""0"" top=""40"" left=""75"" width=""150"" height=""16"" fontsize=""16"" textcolor=""16777215"" fontstyle=""bold""/>"
        
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldTooth"" text=""0"" top=""55"" left=""2"" width=""40"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Button"" name=""IcoGoldTooth"" left=""30"" top=""55"" width=""16"" height=""16"" icon=""25099"" pressedicon=""25099""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkGoldTooth"" left=""55"" top=""55"" width=""20"" height=""20"" text=""""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldToothXP"" text=""0"" top=""55"" left=""75"" width=""150"" height=""16"" textcolor=""16777215"" fontsize=""16"" fontstyle=""bold""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldEgg"" text=""0"" top=""70"" left=""2"" width=""40"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Button"" name=""IcoGoldEgg"" left=""30"" top=""70"" width=""16"" height=""16"" icon=""25098"" pressedicon=""25098""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkGoldEgg"" left=""55"" top=""70"" width=""20"" height=""20"" text=""""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblGoldEggXP"" text=""0"" top=""70"" left=""75"" width=""150"" height=""16"" fontsize=""16"" textcolor=""16777215"" fontstyle=""bold""/>"
       
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblDeepToken"" text=""0"" top=""125"" left=""2"" width=""40"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Button"" name=""IcoDeepToken"" left=""30"" top=""125"" width=""16"" height=""16"" icon=""25107"" pressedicon=""25107""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkDeepToken"" left=""55"" top=""125"" width=""20"" height=""20"" text=""""/>"
   
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnRefreshGolds"" top=""100"" left=""155"" width=""40"" height=""14"" text=""Refresh"" fontsize=""14""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblTotalGolds"" text=""Golds"" top=""100"" left=""2"" width=""100"" height=""16"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblTotalGoldsCount"" text=""0"" top=""100"" left=""75"" width=""40"" height=""16"" fontsize=""16"" textcolor=""255"" fontstyle=""bold""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnDevLine"" left=""0"" top=""115"" width=""215"" height=""5"" textcolor="""" text=""""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblTotalXPlbl"" text=""Total XP"" top=""85"" left=""2"" width=""100"" height=""14"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblTotalXP"" text=""0"" top=""85"" left=""75"" width=""150"" height=""16"" fontsize=""16"" textcolor=""255"" fontstyle=""bold""/>"
      
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.CheckBox"" name=""chkFilterDeepSpam"" left=""55"" top=""140"" width=""20"" height=""20"" text=""""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblFilterDeepSpam"" text=""Filter Deep Spam"" top=""140"" left=""75"" width=""150"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
      
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnSelectATarget"" left=""5"" top=""155"" width=""14"" height=""13"" textcolor=""16777215"" text=""?""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblSelectATarget"" text=""Target Current Selected"" top=""155"" left=""25"" width=""160"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
   
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblCurrentTarget"" text=""Target:"" top=""175"" left=""5"" width=""60"" height=""16"" fontsize=""16"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblCurrentTargetTxt"" text="""" top=""175"" left=""55"" width=""150"" height=""16"" fontsize=""16"" fontstyle=""bold"" textcolor=""255""/>"
   
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.PushButton"" name=""btnStartDeepTrade"" top=""200"" left=""70"" width=""50"" height=""16"" facecolor=""64"" fontsize=""14"" textcolor=""65280"" text=""START""/>"

    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
    
    ViewSchemaOld = ViewSchemaOld & "<page label=""ABOUT"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""FsAbout"" clipped="""">"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName"" text=""THE DEEP"" top=""5"" left=""75"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName2"" text=""For Decal v3 Alpha 7"" top=""18"" left=""40"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutContact"" text=""Created By: Ftuoil Xelrash of Darktide"" top=""35"" left=""5"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutBuildDate"" text=""Plug-in Version: "" top=""50"" left=""30"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""imgLogo"" id=""4567"" left=""60"" top=""75"" width=""80"" height=""120"">"
     
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
       
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</view>"
    
    Exit Function

err_dead:
    Hub.SendText "err_ViewSchemaOld: " & Err.Description, 1
End Function

' Clear the MainView
Private Sub Class_Terminate()
On Error GoTo err_dead

  Set MainView = Nothing
  Exit Sub
  
err_dead:
    Hub.SendText "err_Class_terminate: " & Err.Description, 1
End Sub
