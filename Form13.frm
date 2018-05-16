VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "òÊ—Â"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13065
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "À» "
      Height          =   495
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÃœÊ· „ﬁœ«—Ì"
      Height          =   465
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÃœÊ· »Â«Ì  „«„ ‘œÂ"
      Height          =   465
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÃœÊ· „ﬁœ«—Ì Ê —Ì«·Ì"
      Height          =   465
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   495
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   9840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   1
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   0
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   2
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   3
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   11400
      Top             =   3360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   15901
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· „ﬁœ«—Ì —Ì«·Ì"
      TabPicture(0)   =   "Form13.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÃœÊ· ﬁÌ„   „«„ ‘œÂ Ê«Õœ"
      TabPicture(1)   =   "Form13.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÃœÊ· „ﬁœ«—Ì"
      TabPicture(2)   =   "Form13.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List1"
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).ControlCount=   2
      Begin VB.ListBox List1 
         Height          =   3510
         ItemData        =   "Form13.frx":2D4E
         Left            =   -74880
         List            =   "Form13.frx":2D70
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form13.frx":2DE6
         Height          =   8415
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„"
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
            DataField       =   "ghotr"
            Caption         =   "ﬁÿ—"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
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
         BeginProperty Column04 
            DataField       =   "tolidteydoremeghdar"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â"
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
         BeginProperty Column05 
            DataField       =   "naghlbebadmeghdar"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ"
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
         BeginProperty Column06 
            DataField       =   "zay"
            Caption         =   "÷«Ì⁄« "
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
         BeginProperty Column07 
            DataField       =   "mojodiendmeghdar"
            Caption         =   "„ÊÃÊœÌ «‰ Â«Ì œÊ—Â"
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
               Button          =   -1  'True
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Alignment       =   3
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form13.frx":2DFB
         Height          =   8415
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„"
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
            DataField       =   "ghotr"
            Caption         =   "ﬁÿ—"
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
            DataField       =   "standard8"
            Caption         =   "«” «‰œ«—œ  Ê·Ìœ œ— 8 ”«⁄ "
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
         BeginProperty Column04 
            DataField       =   "mezantolidmostaghim"
            Caption         =   "„Ì“«‰  Ê·Ìœ „” ﬁÌ„"
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
         BeginProperty Column05 
            DataField       =   "zaribtahsimdarsaat"
            Caption         =   "÷—Ì»  Â”Ì„"
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
         BeginProperty Column06 
            DataField       =   "mavadaval"
            Caption         =   "‰ﬁ· «“ Ê«Õœ ﬁ»·"
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
         BeginProperty Column07 
            DataField       =   "zaribdastmozd"
            Caption         =   "÷—Ì» œ” „“œ"
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
         BeginProperty Column08 
            DataField       =   "dastmozd"
            Caption         =   "œ” „“œ"
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
         BeginProperty Column09 
            DataField       =   "zaribsarbar"
            Caption         =   "÷—Ì» ”—»«—"
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
         BeginProperty Column10 
            DataField       =   "sarbar"
            Caption         =   "”—»«—"
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
         BeginProperty Column11 
            DataField       =   "estelak"
            Caption         =   "«” Â·«ò"
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
         BeginProperty Column12 
            DataField       =   "gheymattamam"
            Caption         =   "ﬁÌ„   „«„ ‘œÂ"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
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
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form13.frx":2E10
         Height          =   8415
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„"
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
            DataField       =   "ghotr"
            Caption         =   "ﬁÿ—"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â „ﬁœ«—"
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
         BeginProperty Column04 
            DataField       =   "mojodiavalmemoney"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â „»·€"
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
         BeginProperty Column05 
            DataField       =   "tolidteydoremeghdar"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â „ﬁœ«—"
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
         BeginProperty Column06 
            DataField       =   "tolidteydoremoney"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â „»·€"
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
         BeginProperty Column07 
            DataField       =   "naghlbebadmeghdar"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ „ﬁœ«—"
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
         BeginProperty Column08 
            DataField       =   "naghlbebadmoney"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ „»·€"
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
         BeginProperty Column09 
            DataField       =   "zay"
            Caption         =   "÷«Ì⁄«  „ﬁœ«—"
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
         BeginProperty Column10 
            DataField       =   "zaymoney"
            Caption         =   "÷«Ì⁄«  „»·€"
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
         BeginProperty Column11 
            DataField       =   "mojodiendmeghdar"
            Caption         =   "„ÊÃÊœÌ ¬Œ— œÊ—Â „ﬁœ«—"
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
         BeginProperty Column12 
            DataField       =   "mojodiendmoney"
            Caption         =   "„ÊÃÊœÌ ¬Œ— œÊ—Â „»·€"
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
            BeginProperty Column09 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10920
      Top             =   1440
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
      RecordSource    =   "marahelnameasl"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10920
      Top             =   1200
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
      RecordSource    =   "Koreh"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã„⁄ :"
      Height          =   495
      Index           =   3
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "œ” „“œ"
      Height          =   495
      Index           =   4
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰ﬁ· «“ Ê«Õœ ﬁ»·"
      Height          =   495
      Index           =   0
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "”—»«—"
      Height          =   495
      Index           =   1
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«” Â·«ò"
      Height          =   495
      Index           =   2
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newq As String, newq1 As String, sd As String

Public Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs1(10) As New ADODB.Recordset
Dim rs1number(10) As String
Dim endmefield(20) As String

q = 0
w = 0
db1.Open Form3.Text10.Text
  rs1(0).Open "SELECT * FROM rad ORDER BY rad", db1
    rs1(0).MoveFirst
    Do
      If rs1(0).Fields!Name = "òÊ—Â" Then
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        db2.Open Form3.Text10.Text
          rs1(5).Open "UPDATE Koreh SET [mavadaval]='" + rs1(0).Fields!naghlbebadmoney + "' WHERE (ghotr='" + rs1(0).Fields!ghotr + "')", db2
        db2.Close
      End If
      rs1(0).MoveNext
    Loop Until rs1(0).EOF = True
  rs1(0).Close

  rs1(0).Open "SELECT * FROM sanaveye ORDER BY rad", db1
    rs1(0).MoveFirst
    Do
      If rs1(0).Fields!Name = "òÊ—Â" Then
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        db2.Open Form3.Text10.Text
          rs1(5).Open "UPDATE Koreh SET [mavadaval]='" + rs1(0).Fields!naghlbebadmoney + "' WHERE (ghotr='" + rs1(0).Fields!ghotr + "')", db2
        db2.Close
      End If
      rs1(0).MoveNext
    Loop Until rs1(0).EOF = True
  rs1(0).Close

  rs1(0).Open "SELECT * FROM nahaee ORDER BY rad", db1
    rs1(0).MoveFirst
    Do
      If rs1(0).Fields!Name = "òÊ—Â" Then
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        db2.Open Form3.Text10.Text
          rs1(5).Open "UPDATE Koreh SET [mavadaval]='" + rs1(0).Fields!naghlbebadmoney + "' WHERE (ghotr='" + rs1(0).Fields!ghotr + "')", db2
        db2.Close
      End If
      rs1(0).MoveNext
    Loop Until rs1(0).EOF = True
  rs1(0).Close
db1.Close

Form3.Adodc2.Recordset.Find "name='Koreh'", , adSearchForward, 1
Form26.Adodc4.Recordset.Find "rad=" + Form3.Adodc2.Recordset.Fields!store5, , adSearchForward, 1
Form3.Adodc2.Recordset.Fields!store1 = w
Form3.Adodc2.Recordset.Fields!store2 = Form26.Adodc4.Recordset.Fields!dastmozd
Form3.Adodc2.Recordset.Fields!store3 = Val(Form26.Adodc4.Recordset.Fields!sarbarvahed) + Val(Form26.Adodc4.Recordset.Fields!sarbarjazb)
Form3.Adodc2.Recordset.Fields!store4 = Form26.Adodc4.Recordset.Fields!estehlak
Form3.Adodc2.Recordset.Update

Form13.Text1(0).Text = Form3.Adodc2.Recordset.Fields!store1
Form13.Text1(1).Text = Form3.Adodc2.Recordset.Fields!store2
Form13.Text1(2).Text = Form3.Adodc2.Recordset.Fields!store3
Form13.Text1(3).Text = Form3.Adodc2.Recordset.Fields!store4

Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
newq1 = Adodc1.Recordset.Fields!tolidteydoremeghdar

'Adodc3.Recordset.Find "rad=99997", , adSearchForward, 1
'Adodc3.Recordset.Fields!tolidteydoremeghdar = newq1
'Adodc3.Recordset.Update

Adodc1.Refresh
Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
newq = Adodc1.Recordset.Fields!zaribtahsimdarsaat

Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from Koreh ORDER BY rad ASC"
Adodc3.Refresh

Adodc3.Refresh
Adodc3.Recordset.MoveFirst
Do
   If (Val(Adodc3.Recordset.Fields!rad) <> 99999) And (Val(Adodc3.Recordset.Fields!rad) <> 99997) Then
      'ÃœÊ· „ﬁœ«—Ì
      Form15.Adodc1.RecordSource = "select sum(sheet1number) as amin12 from ozanmain WHERE (gothr='" + Adodc3.Recordset.Fields!ghotr + "') "
      Form15.Adodc1.Refresh
      If (Form15.Adodc1.Recordset.RecordCount = 0) Or IsNull(Form15.Adodc1.Recordset.Fields!amin12) Then
        Adodc3.Recordset.Fields!naghlbebadmeghdar = 0
      Else
        Adodc3.Recordset.Fields!naghlbebadmeghdar = Form15.Adodc1.Recordset.Fields!amin12 - Adodc3.Recordset.Fields!zay
      End If
      q = Adodc3.Recordset.Fields!mojodiavalmeghdar
      w = Adodc3.Recordset.Fields!naghlbebadmeghdar
      e = Adodc3.Recordset.Fields!mojodiendmeghdar
      f = Adodc3.Recordset.Fields!zay
      Adodc3.Recordset.Fields!tolidteydoremeghdar = (Val(e) + Val(w) + Val(f)) - Val(q)
      
      'ÃœÊ· ﬁÌ„   „«„ ‘œÂ Ê«Õœ
      q = Adodc3.Recordset.Fields!mojodiavalmeghdar
      w = Adodc3.Recordset.Fields!mojodiavalmemoney
      e = Adodc3.Recordset.Fields!tolidteydoremeghdar
      Adodc3.Recordset.Fields!tolidteydoremoney = Round(Adodc3.Recordset.Fields!gheymattamam)
      r = Adodc3.Recordset.Fields!tolidteydoremoney
      If ((Val(q) + Val(e)) * Adodc3.Recordset.Fields!naghlbebadmeghdar) <> 0 Then
        sd = Round((Val(w) + Val(r)) / (Val(q) + Val(e)) * Adodc3.Recordset.Fields!naghlbebadmeghdar)
        Adodc3.Recordset.Fields!naghlbebadmoney = sd
      Else
        Adodc3.Recordset.Fields!naghlbebadmoney = 0
      End If
      
      If ((Val(q) + Val(e)) * Adodc3.Recordset.Fields!mojodiendmeghdar) <> 0 Then
        sd1 = Round((Val(w) + Val(r)) / (Val(q) + Val(e)) * Adodc3.Recordset.Fields!mojodiendmeghdar)
        Adodc3.Recordset.Fields!mojodiendmoney = sd1
      Else
        Adodc3.Recordset.Fields!mojodiendmoney = 0
      End If
        
      If ((Val(q) + Val(e)) * Adodc3.Recordset.Fields!zay) <> 0 Then
        sd2 = Round((Val(w) + Val(r)) / (Val(q) + Val(e)) * Adodc3.Recordset.Fields!zay)
        Adodc3.Recordset.Fields!zaymoney = sd2
      Else
        Adodc3.Recordset.Fields!zaymoney = 0
      End If
      
      
      'Adodc3.Recordset.Fields!zaymoney
      
      '÷—Ì»  Â”Ì„
      Adodc3.Recordset.Fields!mezantolidmostaghim = Adodc3.Recordset.Fields!tolidteydoremeghdar
      r = (Adodc3.Recordset.Fields!standard8 * Adodc3.Recordset.Fields!mezantolidmostaghim)
      Adodc3.Recordset.Fields!zaribtahsimdarsaat = Round(r)

      '„Ê«œ «Ê·ÌÂ
      If newq <> 0 Then
        r1 = (Val(Adodc3.Recordset.Fields!zaribdastmozd) * (Val(Text1(1).Text) / newq) * Adodc3.Recordset.Fields!zaribtahsimdarsaat)
        r2 = (Val(Adodc3.Recordset.Fields!zaribsarbar) * (Val(Text1(2).Text) / newq) * Adodc3.Recordset.Fields!zaribtahsimdarsaat)
        r3 = (Val(Adodc3.Recordset.Fields!zaribsarbar) * (Val(Text1(3).Text) / newq) * Adodc3.Recordset.Fields!zaribtahsimdarsaat)
      Else
        r1 = 0
        r2 = 0
        r3 = 0
      End If
      r = Adodc3.Recordset.Fields!mavadaval
      Adodc3.Recordset.Fields!dastmozd = Round(r1)
      Adodc3.Recordset.Fields!sarbar = Round(r2)
      Adodc3.Recordset.Fields!estelak = Round(r3)
      Adodc3.Recordset.Fields!gheymattamam = Round(Val(r) + Val(r1) + Val(r2) + Val(r3))
      Adodc3.Recordset.Update
   End If
   Adodc3.Recordset.MoveNext
Loop Until Adodc3.Recordset.EOF = True


Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "SELECT sum(standard8) as standard81,sum(mezantolidmostaghim) as mezantolidmostaghim1,sum(zaribtahsimdarsaat) as zaribtahsimdarsaat1,sum(mavadaval) as mavadaval1,sum(zaribdastmozd) as zaribdastmozd1,sum(dastmozd) as dastmozd1,sum(zaribsarbar) as zaribsarbar1,sum(sarbar) as sarbar1,sum(estelak) as estelak1,sum(gheymattamam) as gheymattamam1,sum(mojodiavalmeghdar) as mojodiavalmeghdar1, sum(mojodiavalmemoney) as mojodiavalmemoney1,sum(tolidteydoremeghdar) as tolidteydoremeghdar1,sum(tolidteydoremoney) as tolidteydoremoney1,sum(naghlbebadmoney) as naghlbebadmoney1,sum(naghlbebadmeghdar) as naghlbebadmeghdar1,sum(mojodiendmeghdar) as mojodiendmeghdar1,sum(mojodiendmoney) as mojodiendmoney1,sum(zay) as zay1,sum(zaymoney) as zaymoney1 From Koreh WHERE (rad <> 99999) and (rad <> 99998) and (rad <> 99997)"
Adodc3.Refresh

Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
Adodc1.Recordset.Fields!standard8 = Adodc3.Recordset.Fields!standard81
Adodc1.Recordset.Fields!mezantolidmostaghim = Adodc3.Recordset.Fields!mezantolidmostaghim1
Adodc1.Recordset.Fields!zaribtahsimdarsaat = Adodc3.Recordset.Fields!zaribtahsimdarsaat1
Adodc1.Recordset.Fields!mavadaval = Adodc3.Recordset.Fields!mavadaval1
Adodc1.Recordset.Fields!zaribdastmozd = Adodc3.Recordset.Fields!zaribdastmozd1
Adodc1.Recordset.Fields!dastmozd = Adodc3.Recordset.Fields!dastmozd1
Adodc1.Recordset.Fields!zaribsarbar = Adodc3.Recordset.Fields!zaribsarbar1
Adodc1.Recordset.Fields!sarbar = Adodc3.Recordset.Fields!sarbar1
Adodc1.Recordset.Fields!estelak = Adodc3.Recordset.Fields!estelak1
Adodc1.Recordset.Fields!gheymattamam = Adodc3.Recordset.Fields!gheymattamam1
Adodc1.Recordset.Fields!mojodiavalmeghdar = Adodc3.Recordset.Fields!mojodiavalmeghdar1
Adodc1.Recordset.Fields!mojodiavalmemoney = Adodc3.Recordset.Fields!mojodiavalmemoney1
Adodc1.Recordset.Fields!tolidteydoremeghdar = Adodc3.Recordset.Fields!tolidteydoremeghdar1
Adodc1.Recordset.Fields!tolidteydoremoney = Adodc3.Recordset.Fields!tolidteydoremoney1
Adodc1.Recordset.Fields!naghlbebadmoney = Adodc3.Recordset.Fields!naghlbebadmoney1
Adodc1.Recordset.Fields!naghlbebadmeghdar = Adodc3.Recordset.Fields!naghlbebadmeghdar1
Adodc1.Recordset.Fields!mojodiendmeghdar = Adodc3.Recordset.Fields!mojodiendmeghdar1
Adodc1.Recordset.Fields!mojodiendmoney = Adodc3.Recordset.Fields!mojodiendmoney1
Adodc1.Recordset.Fields!zay = Adodc3.Recordset.Fields!zay1
Adodc1.Recordset.Fields!zaymoney = Adodc3.Recordset.Fields!zaymoney1
Adodc1.Recordset.Update
Adodc1.Refresh

Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from Koreh ORDER BY rad ASC"
Adodc3.Refresh


Adodc1.Refresh
DataGrid3.Refresh
DataGrid2.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid3.Refresh
DataGrid2.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid3.Refresh
DataGrid2.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Form33.Label1.Caption = 1
Form33.Show
End Sub

Private Sub Command3_Click()
Form33.Label1.Caption = 2
Form33.Show
End Sub

Private Sub Command4_Click()
Form33.Label1.Caption = 3
Form33.Show
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
Adodc3.Recordset.Update
End Sub

Private Sub DataGrid3_ButtonClick(ByVal ColIndex As Integer)
List1.Visible = True
End Sub

Private Sub Form_Activate()
List1.Visible = False

Form13.Adodc1.ConnectionString = Form3.Text10.Text
Form13.Adodc1.CommandType = adCmdUnknown
Form13.Adodc1.RecordSource = "select * from Koreh ORDER BY rad ASC"
Form13.Adodc1.Refresh

Form3.Adodc2.Recordset.Find "name= 'Koreh'", , adSearchForward, 1
Text1(0).Text = Form3.Adodc2.Recordset.Fields!store1
Text1(1).Text = Form3.Adodc2.Recordset.Fields!store2
Text1(2).Text = Form3.Adodc2.Recordset.Fields!store3
Text1(3).Text = Form3.Adodc2.Recordset.Fields!store4
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub List1_Click()
If (Val(Adodc1.Recordset.Fields!rad) <> 99999) And (Val(Adodc1.Recordset.Fields!rad) <> 99998) And (Val(Adodc1.Recordset.Fields!rad) <> 99997) Then
  DataGrid3.Col = 1
  DataGrid3.Text = List1.Text
End If
List1.Visible = False
End Sub

Private Sub Text1_Change(Index As Integer)
Label1.Caption = Val(Text1(0).Text) + Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
