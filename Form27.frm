VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form27 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÈåÇí ÊãÇã ÔÏå 1"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabHeight       =   520
      TabCaption(0)   =   "ÈåÇí ÊãÇã ÔÏå æÇÍÏ ÑÇäæá ÓÇÒí"
      TabPicture(0)   =   "Form27.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÑÏÔ ÑÇäæá Øí ÏæÑå "
      TabPicture(1)   =   "Form27.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÈåÇí ÊãÇã ÔÏå  ÊæáíÏ æÇÍÏ Óíã æßÇÈá "
      TabPicture(2)   =   "Form27.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ÑÏÔ ãÍÕæá ÏÑ æÇÍÏ Óíã æßÇÈá ÓÇÒí"
      TabPicture(3)   =   "Form27.frx":2D4E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "DataGrid3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "ÈåÇí ÊãÇã ÔÏå ßÇáÇí Óíã æßÇÈá ÝÑæÔ ÑÝÊå "
      TabPicture(4)   =   "Form27.frx":2D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid5"
      Tab(4).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form27.frx":2D86
         Height          =   7335
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "j1sharh"
            Caption         =   "ÔÑÍ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "j1meghdar"
            Caption         =   "ãÞÏÇÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form27.frx":2D9B
         Height          =   7335
         Left            =   -74880
         TabIndex        =   2
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "j2sharh"
            Caption         =   "ÔÑÍ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "j2meghdar1"
            Caption         =   "ãÞÏÇÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "j2meghdar2"
            Caption         =   "ÑíÇá"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form27.frx":2DB0
         Height          =   7335
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "j3sharh"
            Caption         =   "ÔÑÍ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "j3meghdar"
            Caption         =   "ãÞÏÇÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form27.frx":2DC5
         Height          =   7335
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "j4sharh"
            Caption         =   "ÔÑÍ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "j4meghdar"
            Caption         =   "ãÞÏÇÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "Form27.frx":2DDA
         Height          =   7335
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         RightToLeft     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "j5sharh"
            Caption         =   "ÔÑÍ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "j5meghdar"
            Caption         =   "ãÞÏÇÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aminend(5, 10) As String

Private Sub Form_Activate()
Adodc1.Recordset.Find "rad=1", , adSearchForward, 1
'ÌÏæá Çæá
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select sum(mojodipayandoremablagh) as mojodipayandoremablagh1 From g_gardeshmavad Where (nomade=1)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j1meghdar = Adodc2.Recordset.Fields!mojodipayandoremablagh1
Else
  Adodc1.Recordset.Fields!j1meghdar = 0
End If
aminend(1, 1) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc2.RecordSource = "Select sum(moneyonedoremeghdar) as moneyonedoremeghdar1,sum(moneyonedoremablagh) as moneyonedoremablagh1 From g_gardeshmavad Where (nomade=2)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j2meghdar1 = Adodc2.Recordset.Fields!moneyonedoremeghdar1
  Adodc1.Recordset.Fields!j2meghdar2 = Adodc2.Recordset.Fields!moneyonedoremablagh1
Else
  Adodc1.Recordset.Fields!j2meghdar1 = 0
  Adodc1.Recordset.Fields!j2meghdar2 = 0
End If
aminend(2, 1) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 1) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc2.RecordSource = "Select masrafteydoremablagh From ghardeshmavad Where (nomade=1) and (idmade=1)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j3meghdar = Adodc2.Recordset.Fields!masrafteydoremablagh
Else
  Adodc1.Recordset.Fields!j3meghdar = 0
End If
aminend(3, 1) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc2.Recordset.Fields!kaladarjaryanavaldore
Else
  Adodc1.Recordset.Fields!j4meghdar = 0
End If
  
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=13)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc1.Recordset.Fields!j4meghdar - Adodc2.Recordset.Fields!kaladarjaryanavaldore
Else
  Adodc1.Recordset.Fields!j4meghdar = Adodc1.Recordset.Fields!j4meghdar
End If
aminend(4, 1) = Adodc1.Recordset.Fields!j4meghdar

Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=2", , adSearchForward, 1
'ÌÏæá Çæá
aminend(1, 2) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc2.RecordSource = "Select sum(kharidteydoremeghdar) as kharidteydoremeghdar1,sum(kharidteydoremablagh) as kharidteydoremablagh1 From g_gardeshmavad Where (nomade=2)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j2meghdar1 = Adodc2.Recordset.Fields!kharidteydoremeghdar1
  Adodc1.Recordset.Fields!j2meghdar2 = Adodc2.Recordset.Fields!kharidteydoremablagh1
Else
  Adodc1.Recordset.Fields!j2meghdar1 = 0
  Adodc1.Recordset.Fields!j2meghdar2 = 0
End If
aminend(2, 2) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 2) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc2.RecordSource = "Select granol From Exteroder where (rad=99999)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j3meghdar = Adodc2.Recordset.Fields!granol
Else
  Adodc1.Recordset.Fields!j3meghdar = 0
End If
aminend(3, 2) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc2.Recordset.Fields!bahayevahed
Else
  Adodc1.Recordset.Fields!j4meghdar = 0
End If
  
aminend(4, 2) = Adodc1.Recordset.Fields!j4meghdar

Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=3", , adSearchForward, 1
'ÌÏæá Çæá
Adodc1.Recordset.Fields!j1meghdar = Val(aminend(1, 1)) + Val(aminend(1, 2))
aminend(1, 3) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
aminend(2, 3) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 3) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
aminend(3, 3) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc1.Recordset.Fields!j4meghdar = Val(aminend(4, 1)) + Val(aminend(4, 2))
aminend(4, 3) = Adodc1.Recordset.Fields!j4meghdar
  

Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=4", , adSearchForward, 1
'ÌÏæá Çæá
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=998)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j1meghdar = Adodc2.Recordset.Fields!dastmozd
Else
  Adodc1.Recordset.Fields!j1meghdar = 0
End If
aminend(1, 4) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc1.Recordset.Fields!j2meghdar1 = Val(aminend(2, 1)) + Val(aminend(2, 2)) + Val(aminend(2, 3))
Adodc1.Recordset.Fields!j2meghdar2 = Val(aminend(0, 1)) + Val(aminend(0, 2)) + Val(aminend(0, 3))
aminend(2, 4) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 4) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc1.Recordset.Fields!j3meghdar = Val(aminend(3, 1)) + Val(aminend(3, 2)) + Val(aminend(3, 3))
aminend(3, 4) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc2.Recordset.Fields!hazkalapayandore
Else
  Adodc1.Recordset.Fields!j4meghdar = 0
End If
  
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=13)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc1.Recordset.Fields!j4meghdar - Adodc2.Recordset.Fields!hazkalapayandore
Else
  Adodc1.Recordset.Fields!j4meghdar = Adodc1.Recordset.Fields!j4meghdar
End If
aminend(4, 4) = Adodc1.Recordset.Fields!j4meghdar

Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=5", , adSearchForward, 1
'ÌÏæá Çæá
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=998)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j1meghdar = Adodc2.Recordset.Fields!sarbarvahed
Else
  Adodc1.Recordset.Fields!j1meghdar = 0
End If
aminend(1, 5) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc2.RecordSource = "Select sum(foroshteydoremeghdar) as foroshteydoremeghdar1,sum(foroshteydoremablagh) as foroshteydoremablagh1 From g_gardeshmavad Where (nomade=2)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j2meghdar1 = Adodc2.Recordset.Fields!foroshteydoremeghdar1
  Adodc1.Recordset.Fields!j2meghdar2 = Adodc2.Recordset.Fields!foroshteydoremablagh1
Else
  Adodc1.Recordset.Fields!j2meghdar1 = 0
  Adodc1.Recordset.Fields!j2meghdar2 = 0
End If
aminend(2, 5) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 5) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j3meghdar = Adodc2.Recordset.Fields!dastmozd
Else
  Adodc1.Recordset.Fields!j3meghdar = 0
End If
aminend(3, 5) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j4meghdar = Adodc2.Recordset.Fields!zayeat
Else
  Adodc1.Recordset.Fields!j4meghdar = 0
End If
  
aminend(4, 5) = Adodc1.Recordset.Fields!j4meghdar

Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=6", , adSearchForward, 1
'ÌÏæá Çæá
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=998)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j1meghdar = Adodc2.Recordset.Fields!estehlak
Else
  Adodc1.Recordset.Fields!j1meghdar = 0
End If
aminend(1, 6) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc2.RecordSource = "Select sum(zayeatmeghdar) as zayeatmeghdar1,sum(zayeatmablagh) as zayeatmablagh1 From g_gardeshmavad Where (nomade=2)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j2meghdar1 = Adodc2.Recordset.Fields!zayeatmeghdar1
  Adodc1.Recordset.Fields!j2meghdar2 = Adodc2.Recordset.Fields!zayeatmablagh1
Else
  Adodc1.Recordset.Fields!j2meghdar1 = 0
  Adodc1.Recordset.Fields!j2meghdar2 = 0
End If
aminend(2, 6) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 6) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j3meghdar = Val(Adodc2.Recordset.Fields!sarbarvahed) + Val(Adodc2.Recordset.Fields!sarbarjazb)
Else
  Adodc1.Recordset.Fields!j3meghdar = 0
End If
aminend(3, 6) = Adodc1.Recordset.Fields!j3meghdar

'ÌÏæá åÇÑã
Adodc1.Recordset.Fields!j4meghdar = Val(aminend(4, 3)) - Val(aminend(4, 4)) - Val(aminend(4, 5))
aminend(4, 6) = Adodc1.Recordset.Fields!j4meghdar

Adodc1.Recordset.Update



Adodc1.Recordset.Find "rad=7", , adSearchForward, 1
'ÌÏæá Çæá
Adodc1.Recordset.Fields!j1meghdar = Val(aminend(1, 4)) + Val(aminend(1, 5)) + Val(aminend(1, 6))
aminend(1, 7) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc2.RecordSource = "Select sum(mojodipayandoremeghdar) as mojodipayandoremeghdar1,sum(mojodipayandoremablagh) as mojodipayandoremablagh1 From g_gardeshmavad Where (nomade=2)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j2meghdar1 = Adodc2.Recordset.Fields!mojodipayandoremeghdar1
  Adodc1.Recordset.Fields!j2meghdar2 = Adodc2.Recordset.Fields!mojodipayandoremablagh1
Else
  Adodc1.Recordset.Fields!j2meghdar1 = 0
  Adodc1.Recordset.Fields!j2meghdar2 = 0
End If
aminend(2, 7) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 7) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "Select * From sarbar_4 Where (rad=997)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j3meghdar = Val(Adodc2.Recordset.Fields!estehlak)
Else
  Adodc1.Recordset.Fields!j3meghdar = 0
End If
aminend(3, 7) = Adodc1.Recordset.Fields!j3meghdar

Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=8", , adSearchForward, 1
'ÌÏæá Çæá
Adodc1.Recordset.Fields!j1meghdar = Val(aminend(1, 3)) + Val(aminend(1, 7))
aminend(1, 8) = Adodc1.Recordset.Fields!j1meghdar

'ÌÏæá Ïæã
Adodc1.Recordset.Fields!j2meghdar1 = Val(aminend(2, 3)) - (Val(aminend(2, 5)) + Val(aminend(2, 6)) + Val(aminend(2, 7)))
Adodc1.Recordset.Fields!j2meghdar2 = Val(aminend(0, 3)) - (Val(aminend(0, 5)) + Val(aminend(0, 6)) + Val(aminend(0, 7)))
aminend(2, 8) = Adodc1.Recordset.Fields!j2meghdar1
aminend(0, 8) = Adodc1.Recordset.Fields!j2meghdar2

'ÌÏæá Óæã
Adodc1.Recordset.Fields!j3meghdar = Val(aminend(3, 5)) + Val(aminend(3, 6)) + Val(aminend(3, 7))
aminend(3, 8) = Adodc1.Recordset.Fields!j3meghdar

Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=9", , adSearchForward, 1
'ÌÏæá Óæã
Adodc1.Recordset.Fields!j3meghdar = Val(aminend(3, 8)) + Val(aminend(3, 4))
aminend(3, 9) = Adodc1.Recordset.Fields!j3meghdar

Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=1", , adSearchForward, 1
'ÌÏæá äÌã
Adodc2.RecordSource = "Select * From AnbarMahsol Where (rad=99999)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j5meghdar = Adodc2.Recordset.Fields!mojodiavalmemoney
Else
  Adodc1.Recordset.Fields!j5meghdar = 0
End If
aminend(5, 1) = Adodc1.Recordset.Fields!j5meghdar
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=2", , adSearchForward, 1
'ÌÏæá äÌã
Adodc1.Recordset.Fields!j5meghdar = aminend(4, 6)
aminend(5, 2) = Adodc1.Recordset.Fields!j5meghdar
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=3", , adSearchForward, 1
'ÌÏæá äÌã
Adodc1.Recordset.Fields!j5meghdar = Val(aminend(5, 1)) + Val(aminend(5, 2))
aminend(5, 3) = Adodc1.Recordset.Fields!j5meghdar
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=4", , adSearchForward, 1
'ÌÏæá äÌã
Adodc2.RecordSource = "Select * From AnbarMahsol Where (rad=99999)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j5meghdar = Adodc2.Recordset.Fields!mojodiendmoney
Else
  Adodc1.Recordset.Fields!j5meghdar = 0
End If
aminend(5, 4) = Adodc1.Recordset.Fields!j5meghdar
Adodc1.Recordset.Update

Adodc1.Recordset.Find "rad=5", , adSearchForward, 1
'ÌÏæá äÌã
aminend(5, 5) = Adodc1.Recordset.Fields!j5meghdar

Adodc1.Recordset.Find "rad=6", , adSearchForward, 1
'ÌÏæá äÌã
Adodc1.Recordset.Fields!j5meghdar = Val(aminend(5, 3)) + (Val(aminend(5, 4)) + Val(aminend(5, 5)))
aminend(5, 6) = Adodc1.Recordset.Fields!j5meghdar
Adodc1.Recordset.Update


Adodc1.Recordset.Find "rad=8", , adSearchForward, 1
'ÌÏæá äÌã

Adodc2.RecordSource = "Select * From AnbarMahsol Where (rad=99999)"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!j5meghdar = aminend(5, 6) - Adodc2.Recordset.Fields!naghlbebadmoney
Else
  Adodc1.Recordset.Fields!j5meghdar = aminend(5, 6)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

