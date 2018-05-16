VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form24 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” Â·«ﬂ"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "÷—«Ì» ›‰Ì  ”ÂÌ„ «” Â·«ﬂ"
      Height          =   465
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÃœÊ· „Õ«”»Â «” Â·«ﬂ œ«—«∆ÌÂ«Ì À«»  "
      Height          =   465
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· „Õ«”»Â «” Â·«ﬂ œ«—«∆ÌÂ«Ì À«»  "
      TabPicture(0)   =   "Form24.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "÷—«Ì» ›‰Ì  ”ÂÌ„ «” Â·«ﬂ"
      TabPicture(1)   =   "Form24.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form24.frx":2D32
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
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
         ColumnCount     =   21
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
            DataField       =   "name"
            Caption         =   "„—«Õ·  Ê·Ìœ"
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
            DataField       =   "vasayelnaglmarkazi"
            Caption         =   "Ê”«Ì· ‰ﬁ·ÌÂ œ› — „—ﬂ“Ì"
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
            DataField       =   "skarghahi"
            Caption         =   "”«Œ „«‰ ﬂ«—ê«ÂÌ"
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
            DataField       =   "sedari"
            Caption         =   "”«Œ „«‰ «œ«—Ì"
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
            DataField       =   "snaghahban"
            Caption         =   "”«Œ „«‰ ‰êÂ»«‰Ì Ê ⁄„Ê„Ì"
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
         BeginProperty Column06 
            DataField       =   "mashinkarkhane"
            Caption         =   "„«‘Ì‰ ¬·«  ﬂ«—Œ«‰Â (Ê—Êœ «ÿ·«⁄« )"
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
         BeginProperty Column07 
            DataField       =   "tashararat"
            Caption         =   " «”Ì”«  Õ—«— Ì"
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
         BeginProperty Column08 
            DataField       =   "tasab"
            Caption         =   " «”Ì”«  ¬»—”«‰Ì"
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
         BeginProperty Column09 
            DataField       =   "tascool"
            Caption         =   " «”Ì”«  Œ‰ﬂ ﬂ‰‰œÂ"
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
         BeginProperty Column10 
            DataField       =   "tasbargh"
            Caption         =   " «”Ì”«  »—ﬁ —”«‰Ì"
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
         BeginProperty Column11 
            DataField       =   "lavazemazmayeshgah"
            Caption         =   "·Ê«“„ ¬“„«Ì‘ê«ÂÌ"
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
         BeginProperty Column12 
            DataField       =   "asasedari"
            Caption         =   "«À«ÀÌÂ Ê ·Ê«“„ «œ«—Ì"
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
         BeginProperty Column13 
            DataField       =   "vasayelertebati"
            Caption         =   "Ê”«Ì· «— »«ÿÌ"
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
         BeginProperty Column14 
            DataField       =   "mashinsakhteman"
            Caption         =   "„«‘Ì‰ ¬·«  ”«Œ „«‰Ì"
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
         BeginProperty Column15 
            DataField       =   "tasmovaledhava"
            Caption         =   " «”Ì”«  „Ê·œ ÂÊ«"
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
         BeginProperty Column16 
            DataField       =   "vasayelvalefterak"
            Caption         =   "Ê”«Ìÿ ‰ﬁ·ÌÂ Ê ·Ì› —«ﬂ Â«Ì ﬂ«—Œ«‰Â"
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
         BeginProperty Column17 
            DataField       =   "makhazenDOP"
            Caption         =   "„Œ«“‰ Ê  «‰ﬂ Â«Ì DOP"
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
         BeginProperty Column18 
            DataField       =   "tasisgaz"
            Caption         =   " «”Ì”«  ê«“—”«‰Ì"
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
         BeginProperty Column19 
            DataField       =   "abzarkargah"
            Caption         =   "«»“«— ¬·«  ﬂ«—ê«ÂÌ Ê „ ›—ﬁÂ"
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
         BeginProperty Column20 
            DataField       =   "sumend"
            Caption         =   "Ã„⁄ ‰Â«ÌÌ"
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
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form24.frx":2D47
         Height          =   7335
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   9
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
            DataField       =   "name"
            Caption         =   "„—«Õ·  Ê·Ìœ"
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
            DataField       =   "zirbana"
            Caption         =   "“Ì—»‰«Ì"
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
            DataField       =   "abresani"
            Caption         =   "¬» —”«‰Ì"
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
            DataField       =   "cooler"
            Caption         =   "Œ‰ﬂ ﬂ‰‰œÂ"
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
            DataField       =   "naghleye"
            Caption         =   "Ê”«Ìÿ ‰ﬁ·ÌÂ"
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
         BeginProperty Column06 
            DataField       =   "tasisathararati"
            Caption         =   " «”Ì”«  Õ—«— Ì"
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
         BeginProperty Column07 
            DataField       =   "kilovat"
            Caption         =   "»—ﬁ —”«‰Ì"
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
         BeginProperty Column08 
            DataField       =   "asaskarkhane"
            Caption         =   "«À«À ﬂ«—Œ«‰Â"
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
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
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
      RecordSource    =   "Estehlak1"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1440
      Top             =   0
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
      RecordSource    =   "Estehlak2"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2760
      Top             =   0
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp(20) As String

Public Sub Command1_Click()
Dim sd As String
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "SELECT sum(zirbana) as zirbana1,sum(abresani) as abresani1,sum(cooler) as cooler1,sum(naghleye) as naghleye1,sum(tasisathararati) as tasisathararati1,sum(kilovat) as kilovat1,sum(asaskarkhane) as asaskarkhane1 FROM Estehlak1 WHERE (rad <> 999) "
Adodc3.Refresh

Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
Adodc1.Recordset.Fields!zirbana = Adodc3.Recordset.Fields!zirbana1
Adodc1.Recordset.Fields!abresani = Adodc3.Recordset.Fields!abresani1
Adodc1.Recordset.Fields!cooler = Adodc3.Recordset.Fields!cooler1
Adodc1.Recordset.Fields!naghleye = Adodc3.Recordset.Fields!naghleye1
Adodc1.Recordset.Fields!tasisathararati = Adodc3.Recordset.Fields!tasisathararati1
sd = Round(Adodc3.Recordset.Fields!kilovat1, 2)
Adodc1.Recordset.Fields!kilovat = sd
Adodc1.Recordset.Fields!asaskarkhane = Adodc3.Recordset.Fields!asaskarkhane1
Adodc1.Recordset.Update
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Refresh
DataGrid1.Refresh

Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from Estehlak1 ORDER BY rad"
Adodc3.Refresh


  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(0) = Adodc3.Recordset.Fields!zirbana
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(1) = Adodc4.Recordset.Fields!skarghahi
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(2) = Adodc4.Recordset.Fields!tasisgaz
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(3) = Adodc3.Recordset.Fields!tasisathararati
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(4) = Adodc4.Recordset.Fields!tashararat
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(5) = Adodc4.Recordset.Fields!tasab
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(6) = Adodc4.Recordset.Fields!tascool
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(7) = Adodc3.Recordset.Fields!kilovat
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(8) = Adodc4.Recordset.Fields!tasbargh
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(9) = Adodc3.Recordset.Fields!asaskarkhane
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(10) = Adodc4.Recordset.Fields!asasedari
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(11) = Adodc4.Recordset.Fields!vasayelvalefterak
  
  Adodc4.Recordset.MoveFirst
  Do
    If (Adodc4.Recordset.Fields!rad <> 35) And (Adodc4.Recordset.Fields!rad <> 39) And (Adodc4.Recordset.Fields!rad <> 50) And (Adodc4.Recordset.Fields!rad <> 999) Then
    
      Adodc4.Recordset.Fields!vasayelnaglmarkazi = 0
      
      Adodc3.Recordset.Find "rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(1)) / Val(tmp(0))) * Val(Adodc3.Recordset.Fields!zirbana)
      Adodc4.Recordset.Fields!skarghahi = Round(r1)
      
      If Adodc4.Recordset.Fields!rad = 32 Then
        Adodc2.Recordset.Find "rad=50", , adSearchForward, 1
        Adodc4.Recordset.Fields!sedari = Adodc2.Recordset.Fields!sedari
      Else
        Adodc4.Recordset.Fields!sedari = 0
      End If
      
      If Adodc4.Recordset.Fields!rad = 32 Then
        Adodc2.Recordset.Find "rad=50", , adSearchForward, 1
        Adodc4.Recordset.Fields!snaghahban = Adodc2.Recordset.Fields!snaghahban
      Else
        Adodc4.Recordset.Fields!snaghahban = 0
      End If
      
      r1 = (Val(tmp(2)) / Val(tmp(0))) * Val(Adodc3.Recordset.Fields!zirbana)
      Adodc4.Recordset.Fields!tasisgaz = Round(r1)
      
      r1 = (Val(tmp(4)) / Val(tmp(3))) * Val(Adodc3.Recordset.Fields!tasisathararati)
      Adodc4.Recordset.Fields!tashararat = Round(r1)
      
      r1 = ((Val(Adodc3.Recordset.Fields!abresani) / 100) * Val(tmp(5)))
      Adodc4.Recordset.Fields!tasab = Round(r1)
      
      r1 = ((Val(Adodc3.Recordset.Fields!cooler) / 100) * Val(tmp(6)))
      Adodc4.Recordset.Fields!tascool = Round(r1)
      
      r1 = (Val(tmp(8)) / Val(tmp(7))) * Val(Adodc3.Recordset.Fields!kilovat)
      Adodc4.Recordset.Fields!tasbargh = Round(r1)
       
      r1 = (Val(tmp(10)) / Val(tmp(9))) * Val(Adodc3.Recordset.Fields!asaskarkhane)
      Adodc4.Recordset.Fields!asasedari = Round(r1)
      
      If Adodc4.Recordset.Fields!rad = 32 Then
        Adodc2.Recordset.Find "rad=50", , adSearchForward, 1
        Adodc4.Recordset.Fields!vasayelertebati = Adodc2.Recordset.Fields!vasayelertebati
      Else
        Adodc4.Recordset.Fields!vasayelertebati = 0
      End If
      
      If Adodc4.Recordset.Fields!rad = 33 Then
        Adodc2.Recordset.Find "rad=50", , adSearchForward, 1
        Adodc4.Recordset.Fields!tasmovaledhava = Adodc2.Recordset.Fields!tasmovaledhava
      Else
        Adodc4.Recordset.Fields!tasmovaledhava = 0
      End If
      
      r1 = ((Val(Adodc3.Recordset.Fields!naghleye) / 100) * Val(tmp(11)))
      Adodc4.Recordset.Fields!vasayelvalefterak = Round(r1)
      
      If Adodc4.Recordset.Fields!rad = 33 Then
        Adodc2.Recordset.Find "rad=50", , adSearchForward, 1
        Adodc4.Recordset.Fields!abzarkargah = Adodc2.Recordset.Fields!abzarkargah
      Else
        Adodc4.Recordset.Fields!abzarkargah = 0
      End If
      
      Adodc4.Recordset.Fields!sumend = Val(Adodc4.Recordset.Fields!vasayelnaglmarkazi) + Val(Adodc4.Recordset.Fields!skarghahi) + Val(Adodc4.Recordset.Fields!sedari) + Val(Adodc4.Recordset.Fields!snaghahban) + Val(Adodc4.Recordset.Fields!mashinkarkhane) + Val(Adodc4.Recordset.Fields!tashararat) + Val(Adodc4.Recordset.Fields!tasab) + Val(Adodc4.Recordset.Fields!tascool) + Val(Adodc4.Recordset.Fields!tasbargh) + Val(Adodc4.Recordset.Fields!lavazemazmayeshgah) + Val(Adodc4.Recordset.Fields!asasedari) + Val(Adodc4.Recordset.Fields!vasayelertebati) + Val(Adodc4.Recordset.Fields!mashinsakhteman) + Val(Adodc4.Recordset.Fields!tasmovaledhava) + Val(Adodc4.Recordset.Fields!vasayelvalefterak) + Val(Adodc4.Recordset.Fields!makhazenDOP) + Val(Adodc4.Recordset.Fields!tasisgaz) + Val(Adodc4.Recordset.Fields!abzarkargah)
      Adodc4.Recordset.Update
      
    ElseIf (Adodc4.Recordset.Fields!rad = 35) Or (Adodc4.Recordset.Fields!rad = 39) Or (Adodc4.Recordset.Fields!rad = 50) Then
      Adodc4.Recordset.Fields!sumend = Val(Adodc4.Recordset.Fields!vasayelnaglmarkazi) + Val(Adodc4.Recordset.Fields!skarghahi) + Val(Adodc4.Recordset.Fields!sedari) + Val(Adodc4.Recordset.Fields!snaghahban) + Val(Adodc4.Recordset.Fields!mashinkarkhane) + Val(Adodc4.Recordset.Fields!tashararat) + Val(Adodc4.Recordset.Fields!tasab) + Val(Adodc4.Recordset.Fields!tascool) + Val(Adodc4.Recordset.Fields!tasbargh) + Val(Adodc4.Recordset.Fields!lavazemazmayeshgah) + Val(Adodc4.Recordset.Fields!asasedari) + Val(Adodc4.Recordset.Fields!vasayelertebati) + Val(Adodc4.Recordset.Fields!mashinsakhteman) + Val(Adodc4.Recordset.Fields!tasmovaledhava) + Val(Adodc4.Recordset.Fields!vasayelvalefterak) + Val(Adodc4.Recordset.Fields!makhazenDOP) + Val(Adodc4.Recordset.Fields!tasisgaz) + Val(Adodc4.Recordset.Fields!abzarkargah)
      Adodc4.Recordset.Update
    End If
    Adodc4.Recordset.MoveNext
  Loop Until Adodc4.Recordset.EOF = True
  
  Adodc4.CommandType = adCmdUnknown
  Adodc4.RecordSource = "SELECT sum(vasayelnaglmarkazi) as vasayelnaglmarkazi1,sum(skarghahi) as skarghahi1,sum(sedari) as sedari1,sum(snaghahban) as snaghahban1,sum(mashinkarkhane) as mashinkarkhane1,sum(tashararat) as tashararat1,sum(tasab) as tasab1,sum(tascool) as tascool1,sum(tasbargh) as tasbargh1,sum(lavazemazmayeshgah) as lavazemazmayeshgah1,sum(asasedari) as asasedari1,sum(vasayelertebati) as vasayelertebati1,sum(mashinsakhteman) as mashinsakhteman1,sum(tasmovaledhava) as tasmovaledhava1,sum(vasayelvalefterak) as vasayelvalefterak1,sum(makhazenDOP) as makhazenDOP1,sum(tasisgaz) as tasisgaz1,sum(abzarkargah) as abzarkargah1,sum(sumend) as sumend1 FROM Estehlak2 WHERE (rad <> 999) and (rad <> 50) "
  Adodc4.Refresh

  Adodc2.Recordset.Find "rad=999", , adSearchForward, 1
  Adodc2.Recordset.Fields!vasayelnaglmarkazi = Adodc4.Recordset.Fields!vasayelnaglmarkazi1
  Adodc2.Recordset.Fields!skarghahi = Adodc4.Recordset.Fields!skarghahi1
  Adodc2.Recordset.Fields!sedari = Adodc4.Recordset.Fields!sedari1
  Adodc2.Recordset.Fields!snaghahban = Adodc4.Recordset.Fields!snaghahban1
  Adodc2.Recordset.Fields!mashinkarkhane = Adodc4.Recordset.Fields!mashinkarkhane1
  Adodc2.Recordset.Fields!tashararat = Adodc4.Recordset.Fields!tashararat1
  Adodc2.Recordset.Fields!tasab = Adodc4.Recordset.Fields!tasab1
  Adodc2.Recordset.Fields!tascool = Adodc4.Recordset.Fields!tascool1
  Adodc2.Recordset.Fields!tasbargh = Adodc4.Recordset.Fields!tasbargh1
  Adodc2.Recordset.Fields!lavazemazmayeshgah = Adodc4.Recordset.Fields!lavazemazmayeshgah1
  Adodc2.Recordset.Fields!asasedari = Adodc4.Recordset.Fields!asasedari1
  Adodc2.Recordset.Fields!vasayelertebati = Adodc4.Recordset.Fields!vasayelertebati1
  Adodc2.Recordset.Fields!mashinsakhteman = Adodc4.Recordset.Fields!mashinsakhteman1
  Adodc2.Recordset.Fields!tasmovaledhava = Adodc4.Recordset.Fields!tasmovaledhava1
  Adodc2.Recordset.Fields!vasayelvalefterak = Adodc4.Recordset.Fields!vasayelvalefterak1
  Adodc2.Recordset.Fields!makhazenDOP = Adodc4.Recordset.Fields!makhazenDOP1
  Adodc2.Recordset.Fields!tasisgaz = Adodc4.Recordset.Fields!tasisgaz1
  Adodc2.Recordset.Fields!abzarkargah = Adodc4.Recordset.Fields!abzarkargah1
  Adodc2.Recordset.Fields!sumend = Adodc4.Recordset.Fields!sumend1
  Adodc2.Recordset.Update
  Adodc2.Refresh

  Adodc4.CommandType = adCmdUnknown
  Adodc4.RecordSource = "select * from Estehlak2 ORDER BY rad"
  Adodc4.Refresh
  
  Adodc2.Refresh
  DataGrid2.Refresh
  Adodc2.Refresh
  DataGrid2.Refresh
End Sub

Private Sub Command2_Click()
Form50.Label1.Caption = 2
Form50.Show
End Sub

Private Sub Command3_Click()
Form50.Label1.Caption = 1
Form50.Show
End Sub

Private Sub DataGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Adodc2.Recordset.Fields!rad = 50 Then
  blnasd = False
  tmpasd = DataGrid2.Text
Else
  blnasd = True
  tmpasd = DataGrid2.Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub
