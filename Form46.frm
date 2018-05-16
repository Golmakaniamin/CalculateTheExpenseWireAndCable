VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form46 
   Caption         =   "ÇæÒÇä"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "Form46.frx":0000
   LinkTopic       =   "Form46"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   500
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   495
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "Form46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset
Dim Report1 As New ozan
Dim Report2 As New p_allmahsol

If Label1.Caption = 1 Then
  Me.Caption = "ÇæÒÇä"
  Screen.MousePointer = vbHourglass
  CRViewer91.ReportSource = Report1
End If

If Label1.Caption = 2 Then
  Me.Caption = "ÑÏÔ ãÍÕæá"
  db1.Open Form3.Text10.Text
    'Úãá˜ÑÏ ˜ÇáÇ
    rs(0).Open "SELECT Count(idmahsol) As rsnumber FROM amalkardkala WHERE (idmahsol= " + Trim(Str(Form15.Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Form15.Adodc1.Recordset.Fields!rad)) + ")", db1
      If rs(0).Fields!rsnumber > 0 Then
        rs(1).Open "SELECT * FROM amalkardkala WHERE (idmahsol= " + Trim(Str(Form15.Adodc1.Recordset.Fields!idmahsol)) + ") AND (rad =" + Trim(Str(Form15.Adodc1.Recordset.Fields!rad)) + ")", db1
          Report2.Text54.SetText rs(1).Fields!oneanonemojodi
          Report2.Text55.SetText rs(1).Fields!oneantowmojodi
        
          Report2.Text56.SetText rs(1).Fields!oneansummoney
          Report2.Text57.SetText rs(1).Fields!oneansummeter
        
          Report2.Text28.SetText rs(1).Fields!oneanonevaheh
          Report2.Text27.SetText rs(1).Fields!oneantowvahed
        
          Report2.Text30.SetText rs(1).Fields!oneanonemeghdar
          Report2.Text29.SetText rs(1).Fields!oneantowmeghdar
        
          Report2.Text32.SetText rs(1).Fields!seldorevahed
          Report2.Text31.SetText rs(1).Fields!seleghdamvahed
        
          Report2.Text34.SetText rs(1).Fields!seldoremeghdar
          Report2.Text33.SetText rs(1).Fields!seleghdammeghdar
        
          Report2.Text58.SetText rs(1).Fields!seldoremeter
          Report2.Text59.SetText rs(1).Fields!seleghdammeter
        
          Report2.Text60.SetText rs(1).Fields!selcodesum
          Report2.Text61.SetText rs(1).Fields!seleghdamsum
          Report2.Text62.SetText rs(1).Fields!selsum
        
          Report2.Text36.SetText rs(1).Fields!endanonevahed
          Report2.Text35.SetText rs(1).Fields!endantowvahed
        
          Report2.Text38.SetText rs(1).Fields!endanonemeghar
          Report2.Text37.SetText rs(1).Fields!endantowmeghar
        
          Report2.Text63.SetText rs(1).Fields!endanonemeghdar
          Report2.Text64.SetText rs(1).Fields!endantowmeghdar
        
          Report2.Text65.SetText rs(1).Fields!endansum
          
          Report2.Text48.SetText rs(1).Fields!sumtolid
          
        rs(1).Close
      End If
    rs(0).Close
  db1.Close
  
  Report2.Text49.SetText Form15.Combo1.Text
    
  Report2.Text70.SetText Form15.Adodc1.Recordset.Fields!Size
  Report2.Text68.SetText Form15.Adodc1.Recordset.Fields!Size
  Report2.Text70.SetText Form15.Adodc1.Recordset.Fields!gothr
  Report2.Text50.SetText Form15.Adodc1.Recordset.Fields!propertikhas
  Report2.Text69.SetText Form15.Adodc1.Recordset.Fields!nomahsol
  Report2.Text51.SetText Form15.Adodc1.Recordset.Fields!kodemahsol
  Report2.Text72.SetText Form15.Adodc1.Recordset.Fields!ger
  Report2.Text73.SetText Form15.Adodc1.Recordset.Fields!nomes
  Report2.Text75.SetText Form15.Adodc1.Recordset.Fields!sheet1number

  Screen.MousePointer = vbHourglass
  CRViewer91.ReportSource = Report2
End If

CRViewer91.ViewReport
Screen.MousePointer = vbDefault

CRViewer91.Refresh
CRViewer91.Refresh
CRViewer91.Refresh
CRViewer91.Refresh
CRViewer91.Refresh
End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub
