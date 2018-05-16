VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„’—› «” «‰œ«—œ"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ç«Å"
      Height          =   465
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   465
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "»Â«Ì ê—«‰Ê· Ê«Õœ «ò” —Êœ—"
      TabPicture(0)   =   "Form6.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid4"
      Tab(0).Control(1)=   "DataGrid5"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Label4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "„’—› «” «‰œ«—œ"
      TabPicture(1)   =   "Form6.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form6.frx":2D32
         Height          =   4695
         Left            =   3600
         TabIndex        =   5
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         Caption         =   "«” «‰œ«—œ „Ê«œ «Ê·ÌÂ „’—›Ì ÃÂ   Ê·Ìœ Ìﬂ „ —"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "qq"
            Caption         =   "‰«„ „«œÂ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "meghdar2"
            Caption         =   "„ﬁœ«—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
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
         Bindings        =   "Form6.frx":2D47
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         Caption         =   "„Ê«œ „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "qq"
            Caption         =   "‰«„ „«œÂ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0;(0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "meghdar"
            Caption         =   "„ﬁœ«—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form6.frx":2D5C
         Height          =   4695
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         Caption         =   "»Â«Ì ê—«‰Ê·  „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "qq"
            Caption         =   "‰«„ „«œÂ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "meghdar"
            Caption         =   "„ﬁœ«—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0;(0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "meghdar2"
            Caption         =   "ê—œ‘ „Ê«œ"
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
         BeginProperty Column03 
            DataField       =   "sumall"
            Caption         =   "Ê“‰ ò· ê—«‰Ê·"
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
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "Form6.frx":2D71
         Height          =   4695
         Left            =   -71400
         TabIndex        =   13
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         Caption         =   "„Ê«œ „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "qq"
            Caption         =   "‰«„ „«œÂ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0;(0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "meghdar1"
            Caption         =   "„ﬁœ«—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   -71760
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê·Ìœ ÿÌ œÊ—Â :"
         Height          =   495
         Left            =   -69240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê·Ìœ ÿÌ œÊ—Â :"
         Height          =   495
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   5280
         Width           =   2415
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   11880
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Combo3"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ozanmain"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":2D86
      Height          =   5895
      Left            =   7320
      TabIndex        =   2
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
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
      Caption         =   "·Ì”  ò«·« Ê „Ê«œ «Ê·ÌÂ „’—›Ì «” «‰œ«—œ"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "rad"
         Caption         =   "—œÌ›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "kodemahsol"
         Caption         =   "òœ „Õ’Ê·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "nomahsol"
         Caption         =   "‰Ê⁄ „Õ’Ê·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "size"
         Caption         =   "”«Ì“"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "gothr"
         Caption         =   "ﬁÿ—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "propertikhas"
         Caption         =   "„‘Œ’Â Œ«’"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
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
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1440
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ozanunder"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2760
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "masrafestandardmavad2"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4080
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "masrafestandardgranol"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   5400
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   10080
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "p_masrafestandard"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ „Õ’Ê·"
      Height          =   495
      Left            =   13800
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mavadaval(1000) As String, q As Long, w As Long, sumend As String, sum1 As String
Dim adoConnection1 As ADODB.Connection
Dim cmd1 As ADODB.Command
Dim adoRecordset1 As ADODB.Recordset

Private Sub Combo1_Click()
Call Combo1_LostFocus
End Sub

Private Sub Combo1_LostFocus()
If Combo1.ListIndex <> -1 Then
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "select * from ozanmain where idmahsol=" + Combo3.List(Combo1.ListIndex)
  Adodc1.Refresh
'  If DataGrid1.Columns.Count > 6 Then
'    For q = DataGrid1.Columns.Count - 1 To 6 Step -1
'      DataGrid1.Columns.Remove q
'    Next q
'  End If
  
'  List1.Clear
'  Form4.Adodc1.Recordset.MoveFirst5
'  Do
'    List1.AddItem Form4.Adodc1.Recordset.Fields!idmavad
'    DataGrid1.Columns.Add DataGrid1.Columns.Count
'    DataGrid1.Columns.Item(DataGrid1.Columns.Count - 1).Caption = Form4.Adodc1.Recordset.Fields!mavad
'    Form4.Adodc1.Recordset.MoveNext
'  Loop Until Form4.Adodc1.Recordset.EOF = True
'
'  If Adodc1.Recordset.RecordCount > 0 Then
'    Adodc1.Recordset.MoveFirst
'    Do
'      Form15.Adodc2.RecordSource = "SELECT * FROM ozanunder WHERE (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ")"
'      Form15.Adodc2.Refresh
'      If Form15.Adodc2.Recordset.RecordCount > 0 Then
'        DataGrid1.Columns.Item(Form15.Adodc2.Recordset.Fields!idmade + 6).Text = Form15.Adodc2.Recordset.Fields!meghdar
'      End If
'      Adodc1.Recordset.MoveNext
'    Loop Until Adodc1.Recordset.EOF = True
'  End If
End If
End Sub

Public Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs1(10) As New ADODB.Recordset
Dim rs1number(10) As String

db1.Open Form3.Text10.Text
rs1(0).Open "DELETE FROM masrafestandardmavad2", db1
db1.Close

db1.Open Form3.Text10.Text
rs1(0).Open "DELETE FROM masrafestandardgranol", db1
db1.Close

db1.Open Form3.Text10.Text
rs1(0).Open "select count(rad) as rs1number from ozanmain", db1
  rs1number(0) = rs1(0).Fields!rs1number
rs1(0).Close

rs1(0).Open "select * from ozanmain", db1
  If rs1number(0) > 0 Then
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs1number(0)
    ProgressBar1.Value = 0
    rs1(0).MoveFirst
    Do
      ProgressBar1.Value = ProgressBar1.Value + 1
      DoEvents
      
      '«” «‰œ«—œ „Ê«œ «Ê·ÌÂ „’—›Ì ÃÂ   Ê·Ìœ Ìﬂ „ —
      rs1(1).Open "select count(rad) As rs1number from ozanunder where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(1) = rs1(1).Fields!rs1number
      rs1(1).Close
      rs1(1).Open "select * from ozanunder where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ") ORDER BY idmade ASC", db1
      
      ' Ê·Ìœ ÿÌ œÊ—Â «” «‰œ«—œ
      rs1(2).Open "SELECT count(rad) As rs1number FROM amalkardkala WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "SELECT * FROM amalkardkala WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          Label2.Caption = rs1(2).Fields!sumtolid
        Else
          Label2.Caption = 0
        End If

      ' Ê·Ìœ ÿÌ œÊ—Â «ò” —Êœ—
      rs1(3).Open "SELECT count(rad) As rs1number FROM Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(3) = rs1(3).Fields!rs1number
      rs1(3).Close
      rs1(3).Open "SELECT * FROM Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(3) > 0 Then
          Label5.Caption = rs1(3).Fields!zaribtashim
        Else
          Label5.Caption = 0
        End If

      '„Ê«œ „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ
      If rs1(0).Fields!idmahsol = 6 Then
        '”Ì„ ÂÊ«ÌÌ
        Adodc3.ConnectionString = Form3.Text10.Text
        Adodc3.CommandType = adCmdUnknown
        Adodc3.RecordSource = "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ") ORDER BY idmade ASC"
        Adodc3.Refresh
        rs1(5).Open "SELECT SUM(meghdar2) As rssum FROM ozanunder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If IsNull(rs1(5).Fields!rssum) = False Then
          If rs1number(1) > 0 Then
            rs1(1).MoveFirst
            Do
              sd = Round((Label2.Caption / rs1(5).Fields!rssum) * rs1(1).Fields!meghdar2, 4)
              sd1 = Round((Label5.Caption / rs1(5).Fields!rssum) * rs1(1).Fields!meghdar2, 4)
              db2.Open Form3.Text10.Text
                rs1(4).Open "INSERT INTO masrafestandardmavad2 (idmahsol,rad,idmade,meghdar,meghdar1,qq,mainacc) VALUES (" + Trim(Str(rs1(1).Fields!idmahsol)) + "," + Trim(Str(rs1(1).Fields!rad)) + "," + rs1(1).Fields!idmade + "," + Trim(Str(sd)) + "," + Trim(Str(sd1)) + ",'" + rs1(1).Fields!qq + "','" + rs1(1).Fields!mainacc + "')", db2
              db2.Close
              rs1(1).MoveNext
            Loop Until rs1(1).EOF = True
          End If
        End If
        rs1(5).Close
      Else
        Adodc3.ConnectionString = Form3.Text10.Text
        Adodc3.CommandType = adCmdUnknown
        Adodc3.RecordSource = "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ") ORDER BY idmade ASC"
        Adodc3.Refresh
        If rs1number(1) > 0 Then
          rs1(1).MoveFirst
          Do
            sd = Round(Label2.Caption * rs1(1).Fields!meghdar2, 4)
            sd1 = Round(Label5.Caption * rs1(1).Fields!meghdar2, 4)
            db2.Open Form3.Text10.Text
              rs1(4).Open "INSERT INTO masrafestandardmavad2 (idmahsol,rad,idmade,meghdar,meghdar1,qq,mainacc) VALUES (" + Trim(Str(rs1(1).Fields!idmahsol)) + "," + Trim(Str(rs1(1).Fields!rad)) + "," + rs1(1).Fields!idmade + "," + Trim(Str(sd)) + "," + Trim(Str(sd1)) + ",'" + rs1(1).Fields!qq + "','" + rs1(1).Fields!mainacc + "')", db2
            db2.Close
            rs1(1).MoveNext
          Loop Until rs1(1).EOF = True
        End If
      End If
      
    rs1(1).Close
    rs1(2).Close
    rs1(3).Close
    rs1(0).MoveNext
    Loop Until rs1(0).EOF = True

    '»Â«Ì ê—«‰Ê·
    ProgressBar1.Value = 0
    rs1(0).MoveFirst
    Do
      ProgressBar1.Value = ProgressBar1.Value + 1
      DoEvents
    
      '„Ê«œ „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ
      rs1(1).Open "select count(rad) As rs1number from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(1) = rs1(1).Fields!rs1number
      rs1(1).Close
      rs1(1).Open "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ") ORDER BY idmade ASC", db1

      If rs1number(1) > 0 Then
        rs1(1).MoveFirst
        Do
          If (rs1(1).Fields!idmade <> 1) And (rs1(1).Fields!idmade <> 30) Then
            rs1(3).Open "SELECT count(nomade) As rs1number FROM ghardeshmavad WHERE (nomade=1) and (idmade=" + rs1(1).Fields!idmade + ")", db1
              rs1number(3) = rs1(3).Fields!rs1number
            rs1(3).Close
          
            rs1(3).Open "SELECT * FROM ghardeshmavad WHERE (nomade=1) and (idmade=" + rs1(1).Fields!idmade + ")", db1
            If rs1number(3) > 0 Then
              sd = Round(rs1(3).Fields!masrafteydoremablagh)
              sd5 = Trim(Str(sd))
            Else
              sd5 = 0
            End If
          
            rs1(4).Open "SELECT SUM(meghdar1) as amin12 FROM masrafestandardmavad2 WHERE (idmade='" + rs1(1).Fields!idmade + "')", db1
            sd6 = Trim(Str(Val(rs1(4).Fields!amin12)))
            If rs1(4).Fields!amin12 <> 0 Then
              If Val(sd6) <> 0 Then
                sum1 = (Val(sd5) / Val(sd6)) * rs1(1).Fields!meghdar1
              Else
                sum1 = 0
              End If
            Else
              sum1 = 0
            End If
          
            rs1(6).Open "SELECT * FROM infomavad WHERE idmavad=" + rs1(1).Fields!idmade, db1
              If rs1(6).Fields!soz = 0 Then
                db2.Open Form3.Text10.Text
                  rs1(5).Open "INSERT INTO masrafestandardgranol (idmahsol,rad,idmade,meghdar2,meghdar,qq,sumall,mainacc) VALUES (" + Trim(Str(rs1(1).Fields!idmahsol)) + "," + Trim(Str(rs1(1).Fields!rad)) + "," + rs1(1).Fields!idmade + "," + Trim(Str(sd5)) + "," + Trim(Str(sum1)) + ",'" + rs1(1).Fields!qq + "'," + Trim(Str(sd6)) + ",'" + rs1(1).Fields!mainacc + "')", db2
                db2.Close
              Else
            
              End If
            rs1(6).Close
            rs1(3).Close
            rs1(4).Close
          End If
          rs1(1).MoveNext
        Loop Until rs1(1).EOF = True
      End If
    rs1(1).Close
    rs1(0).MoveNext
    Loop Until rs1(0).EOF = True
  End If
rs1(0).Close
db1.Close


'sheet1
db1.Open Form3.Text10.Text
rs1(0).Open "SELECT count(rad1) As rs1number FROM ozanmasir WHERE (rad1=1)", db1
  rs1number(0) = rs1(0).Fields!rs1number
rs1(0).Close
rs1(0).Open "SELECT * FROM ozanmasir WHERE (rad1=1) ", db1
If rs1number(0) > 0 Then
  rs1(0).MoveFirst
  ProgressBar1.Max = rs1number(0)
  ProgressBar1.Value = 0
  Do
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    If (rs1(0).Fields!rad <> 99999) And (rs1(0).Fields!rad <> 99997) Then
      rs1(1).Open "select count(idmade) As rs1number from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ") and (mainacc='*')", db1
        rs1number(1) = rs1(1).Fields!rs1number
      rs1(1).Close
      rs1(1).Open "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ") and (mainacc='*') ORDER BY idmade ASC", db1
        If rs1number(1) > 0 Then
          q = rs1(1).Fields!meghdar
        Else
          q = 0
        End If
      rs1(1).Close
      
      w = 0
      e = 0
      rs1(2).Open "select count(rad) As rs1number from Taab Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Taab Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from Sterander1_6 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Sterander1_6 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from Sterander1_36 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Sterander1_36 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from Sterander1_4 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Sterander1_4 Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from DramToester Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from DramToester Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from Mokhaberat Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Mokhaberat Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close
      
      rs1(2).Open "select count(rad) As rs1number from Exteroder Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        rs1number(2) = rs1(2).Fields!rs1number
      rs1(2).Close
      rs1(2).Open "select * from Exteroder Where (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
        If rs1number(2) > 0 Then
          w = w + rs1(2).Fields!mojodiavalmeghdar
          e = e + rs1(2).Fields!mojodiendmeghdar
        Else
          w = w + 0
          e = e + 0
        End If
      rs1(2).Close

      db2.Open Form3.Text10.Text
        rs1(5).Open "UPDATE ozanmain SET [sheet1number]='" + Trim(Str(Val(e) + Val(q) - Val(w))) + "',[sheet1name]='" + rs1(0).Fields!Name + "' WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db2
      db2.Close
    End If
    rs1(0).MoveNext
  Loop Until rs1(0).EOF = True
End If
rs1(0).Close
db1.Close

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from ozanmain WHERE rad=0"
Adodc1.Refresh

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from ozanunder WHERE rad=0"
Adodc2.Refresh

Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from masrafestandardmavad2 WHERE rad=0"
Adodc3.Refresh

Adodc4.ConnectionString = Form3.Text10.Text
Adodc4.CommandType = adCmdUnknown
Adodc4.RecordSource = "select * from masrafestandardgranol WHERE rad=0"
Adodc4.Refresh

End Sub

Private Sub Command2_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim countrs As String

db1.Open Form3.Text10.Text
rs1.Open "DELETE FROM p_masrafestandard", db1
db1.Close

db1.Open Form3.Text10.Text
rs1.Open "SELECT * FROM ozanmain", db1
rs1.MoveFirst
Do
  rs2.Open "select count(rad) as crad from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ")", db1
  countrs = rs2.Fields!crad
  rs2.Close
  rs2.Open "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(rs1.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ")", db1
  If countrs > 0 Then
    rs2.MoveFirst
    Do
      Adodc6.Refresh
      Adodc6.Recordset.AddNew
      Form2.Adodc1.Recordset.Find "idmahsol=" + Trim(Str(rs1.Fields!idmahsol)), , adSearchForward, 1
      Adodc6.Recordset.Fields!namemahsol = Form2.Adodc1.Recordset.Fields!mahsol
      Adodc6.Recordset.Fields!idmahsol = rs1.Fields!idmahsol
      Adodc6.Recordset.Fields!rad = rs1.Fields!rad
      Adodc6.Recordset.Fields!propertikhas = rs1.Fields!propertikhas
      Adodc6.Recordset.Fields!Size = rs1.Fields!Size
      Adodc6.Recordset.Fields!kodemahsol = rs1.Fields!kodemahsol
      Adodc6.Recordset.Fields!nomahsol = rs1.Fields!nomahsol
      Adodc6.Recordset.Fields!gothr = rs1.Fields!gothr
      Adodc6.Recordset.Fields!idmade = rs2.Fields!idmade
      Adodc6.Recordset.Fields!qq = rs2.Fields!qq
      Adodc6.Recordset.Fields!estmeghdar = rs2.Fields!meghdar
      rs3.Open "select count(rad) as countrs from masrafestandardgranol where (idmahsol=" + Trim(Str(rs2.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs2.Fields!rad)) + ") and (idmade='" + rs2.Fields!idmade + "')", db1
      countrs = rs3.Fields!countrs
      rs3.Close
      
      rs3.Open "select * from masrafestandardgranol where (idmahsol=" + Trim(Str(rs2.Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs2.Fields!rad)) + ") and (idmade='" + rs2.Fields!idmade + "')", db1
      If countrs > 0 Then
        Adodc6.Recordset.Fields!graameghdar = rs3.Fields!meghdar
      Else
        Adodc6.Recordset.Fields!graameghdar = 0
      End If
      rs3.Close
      Adodc6.Recordset.Update
    rs2.MoveNext
    Loop Until rs2.EOF = True
  End If
  rs2.Close
  
  rs1.MoveNext
Loop Until rs1.EOF = True
db1.Close
Form45.Show
End Sub

Private Sub DataGrid1_Click()
Dim sd As String
If Adodc1.Recordset.RecordCount > 0 Then
  '«” «‰œ«—œ „Ê«œ «Ê·ÌÂ „’—›Ì ÃÂ   Ê·Ìœ Ìﬂ „ —
  q = Adodc1.Recordset.Fields!rad
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select * from ozanunder where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ") ORDER BY idmade ASC"
  Adodc2.Refresh
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Form4.Adodc1.Recordset.Find "idmavad=" + Adodc2.Recordset.Fields!idmade, , adSearchForward, 1
      DataGrid2.Col = 0
      DataGrid2.Text = Form4.Adodc1.Recordset.Fields!mavad
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  DataGrid2.Refresh
  
  ' Ê·Ìœ ÿÌ œÊ—Â
  Label2.Caption = 0
  Form5.Adodc2.ConnectionString = Form3.Text10.Text
  Form5.Adodc2.CommandType = adCmdUnknown
  Form5.Adodc2.RecordSource = "SELECT * FROM amalkardkala WHERE (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ")"
  Form5.Adodc2.Refresh
  Label2.Caption = 0
  If Form5.Adodc2.Recordset.RecordCount > 0 Then
    Label2.Caption = Form5.Adodc2.Recordset.Fields!sumtolid
  End If
  
    Label5.Caption = 0
    Form5.Adodc2.ConnectionString = Form3.Text10.Text
    Form5.Adodc2.CommandType = adCmdUnknown
    Form5.Adodc2.RecordSource = "SELECT * FROM Exteroder WHERE (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ")"
    Form5.Adodc2.Refresh
    Label5.Caption = 0
    If Form5.Adodc2.Recordset.RecordCount > 0 Then
      Label5.Caption = Form5.Adodc2.Recordset.Fields!zaribtashim
    End If
  
  '„Ê«œ „’—› ‘œÂ «” «‰œ«—œ ÃÂ   Ê·Ìœ
  Adodc3.ConnectionString = Form3.Text10.Text
  Adodc3.CommandType = adCmdUnknown
  Adodc3.RecordSource = "select * from masrafestandardmavad2 where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ") ORDER BY idmade ASC"
  Adodc3.Refresh
  If Adodc3.Recordset.RecordCount > 0 Then
    Adodc3.Recordset.MoveFirst
    Do
      Form4.Adodc1.Recordset.Find "idmavad=" + Adodc3.Recordset.Fields!idmade, , adSearchForward, 1
      DataGrid3.Col = 0
      DataGrid3.Text = Form4.Adodc1.Recordset.Fields!mavad
      Adodc3.Recordset.MoveNext
    Loop Until Adodc3.Recordset.EOF = True
  End If
  
  DataGrid3.Refresh
  Adodc3.Refresh
  
  DataGrid3.Refresh
  Adodc3.Refresh
  
  '»Â«Ì ê—«‰Ê·
  Adodc4.ConnectionString = Form3.Text10.Text
  Adodc4.CommandType = adCmdUnknown
  Adodc4.RecordSource = "select * from masrafestandardgranol where (idmahsol=" + Trim(Str(Combo3.List(Combo1.ListIndex))) + ") and (rad = " + Trim(Str(q)) + ") ORDER BY idmade ASC"
  Adodc4.Refresh
  
  If Adodc4.Recordset.RecordCount > 0 Then
    Adodc4.Recordset.MoveFirst
    Do
      Form4.Adodc1.Recordset.Find "idmavad=" + Adodc4.Recordset.Fields!idmade, , adSearchForward, 1
      DataGrid4.Col = 0
      DataGrid4.Text = Form4.Adodc1.Recordset.Fields!mavad
      Adodc4.Recordset.MoveNext
    Loop Until Adodc4.Recordset.EOF = True
  End If
  
  DataGrid4.Refresh
  Adodc4.Refresh
  DataGrid4.Refresh
  Adodc4.Refresh
End If
End Sub

Private Sub Form_Activate()
Form2.Adodc1.ConnectionString = Form3.Text10.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "select * from infoMahsol"
Form2.Adodc1.Refresh

If Form2.Adodc1.Recordset.RecordCount > 0 Then
  Combo1.Clear
  Combo3.Clear
  Form2.Adodc1.Recordset.Sort = "idmahsol"
  Form2.Adodc1.Recordset.MoveFirst
  Do
    Combo1.AddItem Form2.Adodc1.Recordset.Fields!mahsol
    Combo3.AddItem Form2.Adodc1.Recordset.Fields!idmahsol
    Form2.Adodc1.Recordset.MoveNext
  Loop Until Form2.Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub MSHFlexGrid1_Click()
MSHFlexGrid2.Row = MSHFlexGrid1.Row
End Sub

