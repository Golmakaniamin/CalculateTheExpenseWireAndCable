VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form48 
   Caption         =   "Form48"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   Icon            =   "Form48.frx":0000
   LinkTopic       =   "Form48"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   8850
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
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim Report1 As New kontrolgardeshmes1
Dim Report2 As New kontrolgardeshmes2
Dim Report3 As New kontrolgardeshmes3

Screen.MousePointer = vbHourglass
Me.Caption = "����� ���� ��"
If Label1.Caption = 1 Then
  CRViewer91.ReportSource = Report1
End If

If Label1.Caption = 2 Then
  CRViewer91.ReportSource = Report2
End If

If Label1.Caption = 3 Then
  CRViewer91.ReportSource = Report3
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


