VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " «»"
   ClientHeight    =   10905
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
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10905
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   10200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÃœÊ· „ﬁœ«—Ì Ê —Ì«·Ì"
      Height          =   465
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÃœÊ· »Â«Ì  „«„ ‘œÂ"
      Height          =   465
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÃœÊ· „ﬁœ«—Ì"
      Height          =   465
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   10320
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   465
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   10320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   2
      Left            =   5280
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   0
      Left            =   9960
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   1
      Left            =   7560
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· „ﬁœ«—Ì —Ì«·Ì"
      TabPicture(0)   =   "Form1.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÃœÊ· ﬁÌ„   „«„ ‘œÂ Ê«Õœ"
      TabPicture(1)   =   "Form1.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÃœÊ· Ê—Êœ «ÿ·«⁄« "
      TabPicture(2)   =   "Form1.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "‰ﬁ· «“ „—Õ·Â ﬁ»·"
      TabPicture(3)   =   "Form1.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid4"
      Tab(3).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form1.frx":2D6A
         Height          =   8415
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
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
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "name"
            Caption         =   "‰«„ „Õ’Ê·"
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
         BeginProperty Column03 
            DataField       =   "gothr"
            Caption         =   "ﬁÿ—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "size_1"
            Caption         =   "”«Ì“"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â „ﬁœ«—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "mojodiendmeghdar"
            Caption         =   "„ÊÃÊœÌ «‰ Â«Ì œÊ—Â „ﬁœ«—"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   3
               Locked          =   -1  'True
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
         Bindings        =   "Form1.frx":2D7F
         Height          =   8415
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   11
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
            Caption         =   "‰«„ „Õ’Ê·"
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
            DataField       =   "kodemahsol"
            Caption         =   "òœ „Õ’Ê·"
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
            DataField       =   "gothr"
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
         BeginProperty Column04 
            DataField       =   "size_1"
            Caption         =   "”«Ì“"
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
            DataField       =   "vaznmes"
            Caption         =   "Ê“‰ „”"
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
            DataField       =   "bahamavad"
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form1.frx":2D94
         Height          =   8415
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            Caption         =   "‰«„ „Õ’Ê·"
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
         BeginProperty Column03 
            DataField       =   "gothr"
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
         BeginProperty Column04 
            DataField       =   "size_1"
            Caption         =   "”«Ì“"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form1.frx":2DA9
         Height          =   8415
         Left            =   -74880
         TabIndex        =   16
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
            DataField       =   "vazn"
            Caption         =   "Ê“‰"
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
            DataField       =   "fey"
            Caption         =   "›Ì"
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
            DataField       =   "mablag"
            Caption         =   "„»·€"
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
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
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
      RecordSource    =   "Taab"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   120
      Top             =   360
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
      RecordSource    =   "Taab1"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰„«Ì‘ ”«Ì“ Â« :"
      Height          =   495
      Index           =   6
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«” Â·«ò"
      Height          =   495
      Index           =   2
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "”—»«—"
      Height          =   495
      Index           =   1
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰ﬁ· «“ Ê«Õœ ﬁ»·"
      Height          =   495
      Index           =   0
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "œ” „“œ"
      Height          =   495
      Index           =   4
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã„⁄ :"
      Height          =   495
      Index           =   3
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newq As String, newq1 As String, sd As String, newbln As Boolean

Private Sub Combo1_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs1(10) As New ADODB.Recordset

If Combo1.ListIndex = 0 Then
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "select * from Taab ORDER BY idmahsol,rad ASC"
  Adodc1.Refresh
  
  db1.Open Form3.Text10.Text
    rs1(0).Open "SELECT sum(vaznmes) as vaznmes1,sum(bahamavad) as bahamavad1,sum(dastmozd) as dastmozd1,sum(sarbar) as sarbar1,sum(estelak) as estelak1,sum(gheymattamam) as gheymattamam1,sum(mojodiavalmeghdar) as mojodiavalmeghdar1, sum(mojodiavalmemoney) as mojodiavalmemoney1,sum(tolidteydoremeghdar) as tolidteydoremeghdar1,sum(tolidteydoremoney) as tolidteydoremoney1,sum(naghlbebadmoney) as naghlbebadmoney1,sum(naghlbebadmeghdar) as naghlbebadmeghdar1,sum(mojodiendmeghdar) as mojodiendmeghdar1,sum(mojodiendmoney) as mojodiendmoney1 FROM Taab WHERE (rad <> 99999)", db1
      db2.Open Form3.Text10.Text
        rs1(5).Open "UPDATE Taab SET [vaznmes]=" + Trim(Str(rs1(0).Fields!vaznmes1)) + ",[bahamavad]=" + Trim(Str(rs1(0).Fields!bahamavad1)) + ",[dastmozd]=" + Trim(Str(rs1(0).Fields!dastmozd1)) + ",[sarbar]=" + Trim(Str(rs1(0).Fields!sarbar1)) + ",[estelak]=" + Trim(Str(rs1(0).Fields!estelak1)) + ",[gheymattamam]=" + Trim(Str(rs1(0).Fields!gheymattamam1)) + ",[mojodiavalmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiavalmeghdar1)) + ",[mojodiavalmemoney]=" + Trim(Str(rs1(0).Fields!mojodiavalmemoney1)) + ",[tolidteydoremeghdar]=" + Trim(Str(rs1(0).Fields!tolidteydoremeghdar1)) + ",[tolidteydoremoney]=" + Trim(Str(rs1(0).Fields!tolidteydoremoney1)) + ",[naghlbebadmoney]=" + Trim(Str(rs1(0).Fields!naghlbebadmoney1)) + ",[naghlbebadmeghdar]=" + Trim(Str(rs1(0).Fields!naghlbebadmeghdar1)) + ",[mojodiendmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiendmeghdar1)) + ",[mojodiendmoney]=" + Trim(Str(rs1(0).Fields!mojodiendmoney1)) + " WHERE (rad=99999)", db2
      db2.Close
    rs1(0).Close
  db1.Close
  
  Adodc1.Refresh
  DataGrid1.Refresh
  DataGrid2.Refresh
  DataGrid3.Refresh
Else
  Adodc1.ConnectionString = Form3.Text10.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM Taab WHERE (gothr='" + Combo1.Text + "') or (rad=99999) ORDER BY idmahsol,rad ASC"
  Adodc1.Refresh
  
  db1.Open Form3.Text10.Text
    rs1(0).Open "SELECT sum(vaznmes) as vaznmes1,sum(bahamavad) as bahamavad1,sum(dastmozd) as dastmozd1,sum(sarbar) as sarbar1,sum(estelak) as estelak1,sum(gheymattamam) as gheymattamam1,sum(mojodiavalmeghdar) as mojodiavalmeghdar1, sum(mojodiavalmemoney) as mojodiavalmemoney1,sum(tolidteydoremeghdar) as tolidteydoremeghdar1,sum(tolidteydoremoney) as tolidteydoremoney1,sum(naghlbebadmoney) as naghlbebadmoney1,sum(naghlbebadmeghdar) as naghlbebadmeghdar1,sum(mojodiendmeghdar) as mojodiendmeghdar1,sum(mojodiendmoney) as mojodiendmoney1 FROM Taab WHERE (gothr='" + Combo1.Text + "') AND (rad <> 99999)", db1
      db2.Open Form3.Text10.Text
        rs1(5).Open "UPDATE Taab SET [vaznmes]=" + Trim(Str(rs1(0).Fields!vaznmes1)) + ",[bahamavad]=" + Trim(Str(rs1(0).Fields!bahamavad1)) + ",[dastmozd]=" + Trim(Str(rs1(0).Fields!dastmozd1)) + ",[sarbar]=" + Trim(Str(rs1(0).Fields!sarbar1)) + ",[estelak]=" + Trim(Str(rs1(0).Fields!estelak1)) + ",[gheymattamam]=" + Trim(Str(rs1(0).Fields!gheymattamam1)) + ",[mojodiavalmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiavalmeghdar1)) + ",[mojodiavalmemoney]=" + Trim(Str(rs1(0).Fields!mojodiavalmemoney1)) + ",[tolidteydoremeghdar]=" + Trim(Str(rs1(0).Fields!tolidteydoremeghdar1)) + ",[tolidteydoremoney]=" + Trim(Str(rs1(0).Fields!tolidteydoremoney1)) + ",[naghlbebadmoney]=" + Trim(Str(rs1(0).Fields!naghlbebadmoney1)) + ",[naghlbebadmeghdar]=" + Trim(Str(rs1(0).Fields!naghlbebadmeghdar1)) + ",[mojodiendmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiendmeghdar1)) + ",[mojodiendmoney]=" + Trim(Str(rs1(0).Fields!mojodiendmoney1)) + " WHERE (rad=99999)", db2
      db2.Close
    rs1(0).Close
  db1.Close
  
  Adodc1.Refresh
  DataGrid1.Refresh
  DataGrid2.Refresh
  DataGrid3.Refresh
End If
End Sub

Public Sub Command1_Click()
On Error Resume Next
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs1(10) As New ADODB.Recordset
Dim rs1number(10) As String
Dim endmefield(20) As String
Dim sd As String

Adodc1.Recordset.Update

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from Taab ORDER BY idmahsol,rad ASC"
Adodc1.Refresh

db2.Open Form3.Text10.Text
  rs1(0).Open "DELETE FROM Taab1", db2
db2.Close

q = 1
w = 0
db1.Open Form3.Text10.Text
  rs1(0).Open "SELECT count(rad) As rs1number FROM rad WHERE Name = ' «»'", db1
     rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "SELECT * FROM rad WHERE Name = ' «»'", db1
      rs1(0).MoveFirst
      Do
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        If rs1(0).Fields!naghlbebadmeghdar <> 0 Then
          sd = Round(rs1(0).Fields!naghlbebadmoney / rs1(0).Fields!naghlbebadmeghdar)
        Else
          sd = 0
        End If
        db2.Open Form3.Text10.Text
          rs1(1).Open "INSERT INTO Taab1 (rad,Name,Name1,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + ",'" + "—«œ" + "','" + "1" + "','" + rs1(0).Fields!ghotr + "','" + rs1(0).Fields!naghlbebadmeghdar + "','" + rs1(0).Fields!naghlbebadmoney + "','" + sd + "')", db2
        db2.Close
        q = q + 1
        rs1(0).MoveNext
      Loop Until rs1(0).EOF = True
    rs1(0).Close
  End If


  rs1(0).Open "SELECT count(rad) As rs1number FROM sanaveye WHERE Name = ' «»'", db1
     rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "SELECT * FROM sanaveye WHERE Name = ' «»'", db1
      rs1(0).MoveFirst
      Do
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        If rs1(0).Fields!naghlbebadmeghdar <> 0 Then
          sd = Round(rs1(0).Fields!naghlbebadmoney / rs1(0).Fields!naghlbebadmeghdar)
        Else
          sd = 0
        End If
        db2.Open Form3.Text10.Text
          rs1(1).Open "INSERT INTO Taab1 (rad,Name,Name1,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + ",'" + "À«‰ÊÌÂ" + "','" + "1" + "','" + rs1(0).Fields!ghotr + "','" + rs1(0).Fields!naghlbebadmeghdar + "','" + rs1(0).Fields!naghlbebadmoney + "','" + sd + "')", db2
        db2.Close
        q = q + 1
        rs1(0).MoveNext
      Loop Until rs1(0).EOF = True
    rs1(0).Close
  End If

  rs1(0).Open "SELECT count(rad) As rs1number FROM nahaee WHERE Name = ' «»'", db1
     rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "SELECT * FROM nahaee WHERE Name = ' «»'", db1
      rs1(0).MoveFirst
      Do
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        If rs1(0).Fields!naghlbebadmeghdar <> 0 Then
          sd = Round(rs1(0).Fields!naghlbebadmoney / rs1(0).Fields!naghlbebadmeghdar)
        Else
          sd = 0
        End If
        db2.Open Form3.Text10.Text
          rs1(1).Open "INSERT INTO Taab1 (rad,Name,Name1,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + ",'" + "‰Â«ÌÌ" + "','" + "1" + "','" + rs1(0).Fields!ghotr + "','" + rs1(0).Fields!naghlbebadmeghdar + "','" + rs1(0).Fields!naghlbebadmoney + "','" + sd + "')", db2
        db2.Close
        q = q + 1
        rs1(0).MoveNext
      Loop Until rs1(0).EOF = True
    rs1(0).Close
  End If

  rs1(0).Open "SELECT count(rad) As rs1number FROM Koreh WHERE Name = ' «»'", db1
     rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "SELECT * FROM Koreh WHERE Name = ' «»'", db1
      rs1(0).MoveFirst
      Do
        w = Val(w) + Val(rs1(0).Fields!naghlbebadmoney)
        If rs1(0).Fields!naghlbebadmeghdar <> 0 Then
          sd = Round(rs1(0).Fields!naghlbebadmoney / rs1(0).Fields!naghlbebadmeghdar)
        Else
          sd = 0
        End If
        db2.Open Form3.Text10.Text
          rs1(1).Open "INSERT INTO Taab1 (rad,Name,Name1,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + ",'" + "òÊ—Â" + "','" + "1" + "','" + rs1(0).Fields!ghotr + "','" + rs1(0).Fields!naghlbebadmeghdar + "','" + rs1(0).Fields!naghlbebadmoney + "','" + sd + "')", db2
        db2.Close
        q = q + 1
        rs1(0).MoveNext
      Loop Until rs1(0).EOF = True
    rs1(0).Close
  End If

Adodc5.ConnectionString = Form3.Text10.Text
Adodc5.CommandType = adCmdUnknown
Adodc5.RecordSource = "SELECT * FROM Taab1 ORDER BY rad"
Adodc5.Refresh
db1.Close

db1.Open Form3.Text10.Text
  rs1(0).Open "SELECT count(rad) As rs1number FROM ozanmasir WHERE name=' «»'", db1
     rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "SELECT * FROM ozanmasir WHERE name=' «»'", db1
       ProgressBar1.Min = 0
       ProgressBar1.Max = rs1number(0)
       ProgressBar1.Value = 0
       Do
         DoEvents
         ProgressBar1.Value = ProgressBar1.Value + 1
         rs1(1).Open "select count(rad) As rs1number from Taab WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
           rs1number(1) = rs1(1).Fields!rs1number
         rs1(1).Close
         
         rs1(1).Open "select * from ozanmain WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
         rs1(2).Open "select * from infoMahsol WHERE idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)), db1
           If rs1number(1) > 0 Then
             db2.Open Form3.Text10.Text
               rs1(5).Open "UPDATE Taab SET [kodemahsol]='" + rs1(1).Fields!kodemahsol + "',[gothr]='" + rs1(1).Fields!gothr + "',[Size_1]='" + rs1(1).Fields!Size + "',[Name]='" + rs1(2).Fields!mahsol + "' WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db2
             db2.Close
           Else
             db2.Open Form3.Text10.Text
               sd2 = rs1(1).Fields!Size
               rs1(5).Open "INSERT INTO Taab (idmahsol,rad,name,kodemahsol,gothr,size_1) VALUES (" + Trim(Str(rs1(1).Fields!idmahsol)) + "," + Trim(Str(rs1(1).Fields!rad)) + ",'" + rs1(2).Fields!mahsol + "','" + rs1(1).Fields!kodemahsol + "','" + rs1(1).Fields!gothr + "','" + sd2 + "')", db2
             db2.Close
           End If
         rs1(1).Close
         rs1(2).Close
         
         rs1(3).Open "SELECT count(rad) As rs1number FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
           rs1number(3) = rs1(3).Fields!rs1number
         rs1(3).Close
         
         If rs1number(3) > 0 Then
           rs1(3).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ") ORDER BY rad1 ASC", db1
             rs1(3).Find "name= ' «»'", , adSearchForward, 1
             p = 0
             o = 0
             If rs1(3).Fields!rad1 > 1 Then
               tmpa = rs1(3).Fields!rad1 - 1
               rs1(3).Find "rad1 =" + Trim(Str(tmpa)), , adSearchForward, 1
               Select Case rs1(3).Fields!Name
                 Case " «»"
                   rs1(4).Open "SELECT count(rad) As rs1number From Taab WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Taab WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = " «»"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If
                 
                 Case "«” —‰œ— 6 +1"
                   rs1(4).Open "SELECT count(rad) As rs1number From Sterander1_6 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Sterander1_6 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "«” —‰œ— 6 +1"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If
                   
                 Case "«” —‰œ— 36 + 1"
                   rs1(4).Open "SELECT count(rad) As rs1number From Sterander1_36 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Sterander1_36 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "«” —‰œ— 36 + 1"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If

                 Case "«” —‰œ— 4 + 1"
                   rs1(4).Open "SELECT count(rad) As rs1number From Sterander1_4 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Sterander1_4 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "«” —‰œ— 4 + 1"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If

                 Case "œ—«„  ÊÌ” —"
                   rs1(4).Open "SELECT count(rad) As rs1number From DramToester WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From DramToester WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "œ—«„  ÊÌ” —"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If

                 Case "„Œ«»—« Ì"
                   rs1(4).Open "SELECT count(rad) As rs1number From Mokhaberat WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Mokhaberat WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "„Œ«»—« Ì"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If

                 Case "«ò” —Êœ—"
                   rs1(4).Open "SELECT count(rad) As rs1number From Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "„Œ«»—« Ì"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If
            
                 Case "»«‰ç—"
                   rs1(4).Open "SELECT count(rad) As rs1number From Bancher WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                      rs1number(4) = rs1(4).Fields!rs1number
                   rs1(4).Close
                   If rs1number(4) > 0 Then
                     rs1(4).Open "SELECT * From Bancher WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
                       w = Val(w) + Val(rs1(4).Fields!naghlbebadmoney)
                       p = rs1(4).Fields!naghlbebadmeghdar
                       o = rs1(4).Fields!naghlbebadmoney
'                       sd = Round(rs1(4).Fields!naghlbebadmoney / rs1(4).Fields!naghlbebadmeghdar)
'                       sd1 = "»«‰ç—"
'                       db2.Open Form3.Text10.Text
'                         rs1(5).Open "INSERT INTO Taab1 (rad,Name,ghotr,vazn,mablag,fey) VALUES (" + Trim(Str(q)) + "," + sd1 + "," + Trim(Str(rs1(3).Fields!ghotr)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmeghdar)) + "," + Trim(Str(rs1(3).Fields!naghlbebadmoney)) + "," + Trim(Str(sd)) + ")", db2
'                       db2.Close
                       q = q + 1
                     rs1(4).Close
                   End If
              End Select
              db2.Open Form3.Text10.Text
                rs1(5).Open "UPDATE Taab SET [exist]='1' ,[tolidteydoremeghdar]='" + Trim(Str(p)) + "',[bahamavad]=" + Trim(Str(o)) + " WHERE (idmahsol=" + Trim(Str(rs1(3).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(3).Fields!rad)) + ")", db2
              db2.Close
             Else
              db2.Open Form3.Text10.Text
                rs1(5).Open "UPDATE Taab SET [exist]='0' WHERE (idmahsol=" + Trim(Str(rs1(3).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(3).Fields!rad)) + ")", db2
              db2.Close
             End If
           rs1(3).Close
         End If
         rs1(0).MoveNext
       Loop Until rs1(0).EOF = True
    rs1(0).Close
  End If
db1.Close


'çÂ«— ›Ì·œ «’·Ì

db2.Open Form3.Text10.Text
  rs1(0).Open "SELECT * FROM sarbar_4 WHERE rad =5", db2
  rs1(5).Open "UPDATE marahelnameasl SET [store1]='0',[store2] ='0',[store3] ='0',[store4] ='0'  WHERE name= 'Taab'", db2
  rs1(6).Open "UPDATE marahelnameasl SET [store1]='" + Trim(Str(w)) + "',[store2] ='" + rs1(0).Fields!dastmozd + "',[store3] ='" + Trim(Str(Val(rs1(0).Fields!sarbarvahed) + Val(rs1(0).Fields!sarbarjazb))) + "',[store4] ='" + rs1(0).Fields!estehlak + "'  WHERE name= 'Taab'", db2
db2.Close

db1.Open Form3.Text10.Text
  rs1(3).Open "SELECT * FROM  marahelnameasl WHERE name= 'Taab'", db1
  rs1(0).Open "select count(rad) As rs1number from Taab", db1
    rs1number(0) = rs1(0).Fields!rs1number
  rs1(0).Close
  If rs1number(0) > 0 Then
    rs1(0).Open "select * from Taab ORDER BY idmahsol,rad ASC", db1
      rs1(0).Find "rad=99999", , adSearchForward, 1
      newq = rs1(0).Fields!vaznmes
      ProgressBar1.Min = 0
      ProgressBar1.Max = rs1number(0)
      ProgressBar1.Value = 0
      rs1(0).MoveFirst
      Do
         DoEvents
         If ProgressBar1.Value < ProgressBar1.Max Then ProgressBar1.Value = ProgressBar1.Value + 1
         If (Val(rs1(0).Fields!rad) <> 99999) Then
           For intcount = 0 To 20
             endmefield(intcount) = ""
           Next intcount
           
           'ÃœÊ· „ﬁœ«—Ì
           If rs1(0).Fields!exist = 0 Then
'             If rs1(0).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
               sd1 = " «»"
               rs1(1).Open "select count(rad) As rs1number from ozanmain WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ") and (sheet1name='" + sd1 + "')", db1
                 rs1number(1) = rs1(1).Fields!rs1number
               rs1(1).Close
               If rs1number(1) > 0 Then
                 rs1(1).Open "select * from ozanmain WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ") and (sheet1name='" + sd1 + "')", db1
                   endmefield(0) = rs1(1).Fields!sheet1number
                 rs1(1).Close
               End If
'             Else
'               MsgBox 1
'             End If
           Else
             endmefield(0) = rs1(0).Fields!tolidteydoremeghdar
           End If
           
           endmefield(1) = (Val(rs1(0).Fields!mojodiavalmeghdar) + Val(endmefield(0))) - Val(rs1(0).Fields!mojodiendmeghdar)

           'ÃœÊ· ﬁÌ„   „«„ ‘œÂ Ê«Õœ
           If rs1(0).Fields!exist = 0 Then
             If (rs1(0).Fields!gothr <> 0) <> 0 Then
               rs1(2).Open "SELECT count(rad) As rs1number FROM Taab1 WHERE (ghotr='" + rs1(0).Fields!gothr + "')", db1
                 rs1number(2) = rs1(2).Fields!rs1number
               rs1(2).Close
               If (rs1number(2) > 0) Then
                 rs1(2).Open "SELECT * FROM Taab1 WHERE (ghotr='" + rs1(0).Fields!gothr + "')", db1
                   If rs1(2).Fields!vazn <> 0 Then
                     endmefield(2) = Round((rs1(2).Fields!mablag / rs1(2).Fields!vazn) * rs1(0).Fields!vaznmes)
                   Else
                     endmefield(2) = 0
                   End If
                 rs1(2).Close
               Else
                 endmefield(2) = 0
               End If
             Else
               endmefield(2) = 0
             End If
           Else
             endmefield(2) = rs1(0).Fields!bahamavad
           End If
      
           If newq <> 0 Then
             r1 = ((Val(rs1(3).Fields!store2) / newq) * endmefield(0))
             r2 = ((Val(rs1(3).Fields!store3) / newq) * endmefield(0))
             r3 = ((Val(rs1(3).Fields!store4) / newq) * endmefield(0))
             endmefield(3) = 0
             endmefield(4) = Round(r1)
             endmefield(5) = Round(r2)
             endmefield(6) = Round(r3)
           Else
             endmefield(3) = 0
             endmefield(4) = 0
             endmefield(5) = 0
             endmefield(6) = 0
           End If
           
           endmefield(7) = Round(Val(endmefield(6)) + Val(endmefield(5)) + Val(endmefield(4)) + Val(endmefield(2)))
           
           'ÃœÊ· —Ì«·Ì
           If ((Val(rs1(0).Fields!mojodiavalmeghdar) + Val(endmefield(0))) * Val(endmefield(1))) <> 0 Then
             sd = ((Val(rs1(0).Fields!mojodiavalmemoney) + Val(endmefield(7))) / ((Val(rs1(0).Fields!mojodiavalmeghdar) + Val(endmefield(0)))) * Val(endmefield(1)))
             endmefield(8) = Round(sd)
           Else
             endmefield(8) = 0
           End If
           
           endmefield(9) = Val(rs1(0).Fields!mojodiavalmemoney) + Val(endmefield(7)) - Val(endmefield(8))
           
           db2.Open Form3.Text10.Text
             tmp1 = endmefield(0)
             tmp2 = endmefield(7)
             rs1(5).Open "UPDATE Taab SET [tolidteydoremeghdar]=" + endmefield(0) + ",[vaznmes]=" + tmp1 + ",[naghlbebadmeghdar]=" + endmefield(1) + ",[bahamavad]=" + endmefield(2) + ",[dastmozd]=" + endmefield(4) + ",[sarbar]=" + endmefield(5) + ",[estelak]=" + endmefield(6) + ",[gheymattamam]=" + endmefield(7) + ",[tolidteydoremoney]=" + tmp2 + ",[naghlbebadmoney]=" + endmefield(8) + ",[mojodiendmoney]=" + endmefield(9) + " WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad = " + Trim(Str(rs1(0).Fields!rad)) + ")", db2
           db2.Close
        End If
        rs1(0).MoveNext
     Loop Until rs1(0).EOF = True
    rs1(3).Close
    rs1(0).Close
  End If
db1.Close

db1.Open Form3.Text10.Text
  rs1(0).Open "SELECT sum(vaznmes) as vaznmes1,sum(bahamavad) as bahamavad1,sum(dastmozd) as dastmozd1,sum(sarbar) as sarbar1,sum(estelak) as estelak1,sum(gheymattamam) as gheymattamam1,sum(mojodiavalmeghdar) as mojodiavalmeghdar1, sum(mojodiavalmemoney) as mojodiavalmemoney1,sum(tolidteydoremeghdar) as tolidteydoremeghdar1,sum(tolidteydoremoney) as tolidteydoremoney1,sum(naghlbebadmoney) as naghlbebadmoney1,sum(naghlbebadmeghdar) as naghlbebadmeghdar1,sum(mojodiendmeghdar) as mojodiendmeghdar1,sum(mojodiendmoney) as mojodiendmoney1 FROM Taab WHERE (rad <> 99999)", db1
    db2.Open Form3.Text10.Text
      rs1(5).Open "UPDATE Taab SET [vaznmes]=" + Trim(Str(rs1(0).Fields!vaznmes1)) + ",[bahamavad]=" + Trim(Str(rs1(0).Fields!bahamavad1)) + ",[dastmozd]=" + Trim(Str(rs1(0).Fields!dastmozd1)) + ",[sarbar]=" + Trim(Str(rs1(0).Fields!sarbar1)) + ",[estelak]=" + Trim(Str(rs1(0).Fields!estelak1)) + ",[gheymattamam]=" + Trim(Str(rs1(0).Fields!gheymattamam1)) + ",[mojodiavalmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiavalmeghdar1)) + ",[mojodiavalmemoney]=" + Trim(Str(rs1(0).Fields!mojodiavalmemoney1)) + ",[tolidteydoremeghdar]=" + Trim(Str(rs1(0).Fields!tolidteydoremeghdar1)) + ",[tolidteydoremoney]=" + Trim(Str(rs1(0).Fields!tolidteydoremoney1)) + ",[naghlbebadmoney]=" + Trim(Str(rs1(0).Fields!naghlbebadmoney1)) + ",[naghlbebadmeghdar]=" + Trim(Str(rs1(0).Fields!naghlbebadmeghdar1)) + ",[mojodiendmeghdar]=" + Trim(Str(rs1(0).Fields!mojodiendmeghdar1)) + ",[mojodiendmoney]=" + Trim(Str(rs1(0).Fields!mojodiendmoney1)) + " WHERE (rad=99999)", db2
    db2.Close
  rs1(0).Close
db1.Close

Adodc1.Refresh
DataGrid3.Refresh
DataGrid2.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Form34.Label1.Caption = 1
Form34.Show
End Sub

Private Sub Command3_Click()
Form34.Label1.Caption = 2
Form34.Show
End Sub

Private Sub Command4_Click()
Form34.Label1.Caption = 3
Form34.Show
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim sorta As String
If Adodc1.Recordset.RecordCount > 0 Then

  Select Case ColIndex
    Case 0
      sorta = "rad"

    Case 1
      sorta = "idmahsol"

    Case 2
      sorta = "kodemahsol"

    Case 3
      sorta = "gothr"

    Case 4
      sorta = "size_1"

    Case 5
      sorta = "vaznmes"

    Case 6
      sorta = "bahamavad"

    Case 7
      sorta = "dastmozd"

    Case 8
      sorta = "sarbar"

    Case 9
      sorta = "estelak"

    Case 10
      sorta = "gheymattamam"

  End Select
  Adodc1.Recordset.Sort = sorta
End If
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
Dim sorta As String
If Adodc1.Recordset.RecordCount > 0 Then

  Select Case ColIndex
    Case 0
      sorta = "rad"

    Case 1
      sorta = "idmahsol"

    Case 2
      sorta = "kodemahsol"

    Case 3
      sorta = "gothr"

    Case 4
      sorta = "size_1"

    Case 5
      sorta = "mojodiavalmeghdar"

    Case 6
      sorta = "mojodiavalmemoney"

    Case 7
      sorta = "tolidteydoremeghdar"

    Case 8
      sorta = "tolidteydoremoney"

    Case 9
      sorta = "naghlbebadmeghdar"

    Case 10
      sorta = "naghlbebadmoney"

    Case 11
      sorta = "mojodiendmeghdar"

    Case 12
      sorta = "mojodiendmoney"

  End Select
  Adodc1.Recordset.Sort = sorta
End If
End Sub

Private Sub DataGrid3_HeadClick(ByVal ColIndex As Integer)
Dim sorta As String
If Adodc1.Recordset.RecordCount > 0 Then

  Select Case ColIndex
    Case 0
      sorta = "rad"

    Case 1
      sorta = "idmahsol"

    Case 2
      sorta = "kodemahsol"

    Case 3
      sorta = "gothr"

    Case 4
      sorta = "size_1"

    Case 5
      sorta = "mojodiavalmeghdar"

    Case 6
      sorta = "mojodiavalmemoney"

    Case 7
      sorta = "mojodiendmeghdar"

  End Select
  Adodc1.Recordset.Sort = sorta
End If
End Sub

Private Sub Form_Activate()
Dim db1 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset

Form1.Adodc1.ConnectionString = Form3.Text10.Text
Form1.Adodc1.CommandType = adCmdUnknown
Form1.Adodc1.RecordSource = "select * from Taab ORDER BY idmahsol,rad ASC"
Form1.Adodc1.Refresh

Form3.Adodc2.Recordset.Find "name= 'Taab'", , adSearchForward, 1
Text1(0).Text = Form3.Adodc2.Recordset.Fields!store1
Text1(1).Text = Form3.Adodc2.Recordset.Fields!store2
Text1(2).Text = Form3.Adodc2.Recordset.Fields!store3
Text1(3).Text = Form3.Adodc2.Recordset.Fields!store4

Label1.Caption = Val(Text1(0).Text) + Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text)

Combo1.Clear
Combo1.AddItem " „«„ „Õ’Ê·« "
db1.Open Form3.Text10.Text
  rs(0).Open "SELECT DISTINCT gothr FROM Taab WHERE (rad <> 99999) ORDER BY gothr", db1
    rs(0).MoveFirst
    Do
      Combo1.AddItem rs(0).Fields!gothr
      rs(0).MoveNext
    Loop Until rs(0).EOF = True
  rs(0).Close
db1.Close
Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

