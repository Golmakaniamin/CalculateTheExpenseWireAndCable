VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form23 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ò‰ —· ê—œ‘ „”"
   ClientHeight    =   9330
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
   Icon            =   "Form23.frx":0000
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· ”Â"
      TabPicture(0)   =   "Form23.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ÃœÊ· œÊ"
      TabPicture(1)   =   "Form23.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command3"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "DataGrid2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ÃœÊ· Ìò"
      TabPicture(2)   =   "Form23.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2"
      Tab(2).Control(1)=   "DataGrid1"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command5 
         Caption         =   "ò‰ —· „’—› ê—«‰Ê·"
         Height          =   465
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   8040
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÃœÊ· ”Â"
         Height          =   465
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   8040
         Width           =   4935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "—Ì«·"
         Height          =   465
         Left            =   -68760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   8040
         Width           =   5895
      End
      Begin VB.CommandButton Command4 
         Caption         =   "„ﬁœ«—"
         Height          =   465
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   8040
         Width           =   5895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ç«Å"
         Height          =   465
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   8040
         Width           =   12015
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form23.frx":2D4E
         Height          =   7455
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13150
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
         ColumnCount     =   7
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
            Caption         =   "‰«„ „—Õ·Â"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
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
            DataField       =   "varedeteydoremeghdar"
            Caption         =   "Ê«—œÂ ÿÌ œÊ—Â"
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
            DataField       =   "enteghalbade"
            Caption         =   "«‰ ﬁ«· »Â Ê«Õœ »⁄œ"
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
            DataField       =   "zayeat"
            Caption         =   "÷«Ì⁄«   Ê·Ìœ"
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
            DataField       =   "mojodienddore"
            Caption         =   "„ÊÃÊœÌ Å«Ì«‰ œÊ—Â"
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form23.frx":2D63
         Height          =   7455
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13150
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
         ColumnCount     =   30
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
            Caption         =   "‰«„ „—Õ·Â"
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
            DataField       =   "sanaveye1"
            Caption         =   "À«‰ÊÌÂ („ﬁœ«—)"
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
            DataField       =   "sanaveye2"
            Caption         =   "À«‰ÊÌÂ (—Ì«·)"
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
            DataField       =   "nahaee1"
            Caption         =   "‰Â«ÌÌ („ﬁœ«—)"
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
            DataField       =   "nahaee2"
            Caption         =   "‰Â«ÌÌ (—Ì«·)"
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
            DataField       =   "Koreh1"
            Caption         =   "òÊ—Â („ﬁœ«—)"
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
            DataField       =   "Koreh2"
            Caption         =   "òÊ—Â (—Ì«·)"
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
            DataField       =   "Taab1"
            Caption         =   " «» („ﬁœ«—)"
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
            DataField       =   "Taab2"
            Caption         =   " «» (—Ì«·)"
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
            DataField       =   "Sterander1_61"
            Caption         =   "«” —‰œ— 6+1 („ﬁœ«—)"
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
            DataField       =   "Sterander1_62"
            Caption         =   "«” —‰œ— 6+1 (—Ì«·)"
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
            DataField       =   "Sterander1_361"
            Caption         =   "«” —‰œ— 36+1 (—Ì«·)"
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
            DataField       =   "Sterander1_362"
            Caption         =   "«” —‰œ— 36+1 („ﬁœ«—)"
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
            DataField       =   "Sterander1_41"
            Caption         =   "«” —‰œ— 4+1 („ﬁœ«—)"
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
            DataField       =   "Sterander1_42"
            Caption         =   "«” —‰œ— 4+1 (—Ì«·)"
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
            DataField       =   "DramToester1"
            Caption         =   "œ—«„  ÊÌ” — („ﬁœ«—)"
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
            DataField       =   "DramToester2"
            Caption         =   "œ—«„  ÊÌ” — (—Ì«·)"
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
            DataField       =   "Mokhaberat1"
            Caption         =   "„Œ«»—«  („ﬁœ«—)"
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
            DataField       =   "Mokhaberat2"
            Caption         =   "„Œ«»—«  (—Ì«·)"
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
            DataField       =   "Exteroder1"
            Caption         =   "«ò” —Êœ— („ﬁœ«—)"
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
         BeginProperty Column21 
            DataField       =   "Exteroder2"
            Caption         =   "«ò” —Êœ— (—Ì«·)"
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
         BeginProperty Column22 
            DataField       =   "Bastebandi1"
            Caption         =   "»” Â »‰œÌ („ﬁœ«—)"
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
         BeginProperty Column23 
            DataField       =   "Bastebandi2"
            Caption         =   "»” Â »‰œÌ (—Ì«·)"
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
         BeginProperty Column24 
            DataField       =   "AnbarMahsol1"
            Caption         =   "«‰»«— „Õ’Ê· („ﬁœ«—)"
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
         BeginProperty Column25 
            DataField       =   "AnbarMahsol2"
            Caption         =   "«‰»«— „Õ’Ê· (—Ì«·)"
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
         BeginProperty Column26 
            DataField       =   "bancher1"
            Caption         =   "»«‰ç—(„ﬁœ«—)"
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
         BeginProperty Column27 
            DataField       =   "bancher2"
            Caption         =   "»«‰ç— (—Ì«·)"
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
         BeginProperty Column28 
            DataField       =   "sum1"
            Caption         =   "Ã„⁄ („ﬁœ«—)"
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
         BeginProperty Column29 
            DataField       =   "sum2"
            Caption         =   "Ã„⁄ (—Ì«·)"
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
            BeginProperty Column21 
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
            BeginProperty Column26 
            EndProperty
            BeginProperty Column27 
            EndProperty
            BeginProperty Column28 
            EndProperty
            BeginProperty Column29 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form23.frx":2D78
         Height          =   7455
         Left            =   7200
         TabIndex        =   4
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   13150
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
            DataField       =   "name"
            Caption         =   ""
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
            DataField       =   "meghdar"
            Caption         =   ""
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
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form23.frx":2D8D
         Height          =   7455
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   13150
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "idmavad"
            Caption         =   "òœ „«œÂ"
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
            DataField       =   "mavad"
            Caption         =   "‰«„ „«œÂ"
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
            DataField       =   "mvage"
            Caption         =   "„’—› Ê«ﬁ⁄Ì"
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
            DataField       =   "mesta"
            Caption         =   "„’—› «” «‰œ«—œ"
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
            DataField       =   "enheraf"
            Caption         =   "«‰Õ—«› „’—› «“ «” «‰œ«—œ"
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
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
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
      RecordSource    =   "kontrolgardeshmes"
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
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "kontrolgardeshmes"
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
      Left            =   2880
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "Ghardeshmes1"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4320
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "infomavad"
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
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim q(10) As String
Dim q1(11) As String

Private Sub Command1_Click()
Form52.Show
End Sub

Private Sub Command2_Click()
Form48.Label1.Caption = 1
Form48.Show
End Sub

Private Sub Command3_Click()
Form48.Label1.Caption = 2
Form48.Show
End Sub

Private Sub Command4_Click()
Form48.Label1.Caption = 3
Form48.Show
End Sub

Private Sub Command5_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

db1.Open Form3.Text10.Text
  rs1.Open "DELETE FROM p_gardeshmavad", db1
db1.Close

Adodc4.ConnectionString = Form3.Text10.Text
Adodc4.CommandType = adCmdUnknown
Adodc4.RecordSource = "SELECT * FROM p_gardeshmavad"
Adodc4.Refresh

db1.Open Form3.Text10.Text
  rs1.Open "SELECT * FROM ghardeshmavad WHERE nogra='1'", db1
    rs1.MoveFirst
    Do
        Adodc4.Refresh
        Adodc4.Recordset.AddNew
        Adodc4.Recordset.Fields!idmade = rs1.Fields!idmade
        Adodc4.Recordset.Fields!nomade = rs1.Fields!nomade
        
        Adodc4.Recordset.Fields!moneyonedoremeghdar = rs1.Fields!moneyonedoremeghdar
        Adodc4.Recordset.Fields!kharidteydoremeghdar = rs1.Fields!kharidteydoremeghdar
        Adodc4.Recordset.Fields!naghlazgeranolmeghdar = rs1.Fields!naghlazgeranolmeghdar
        
        Adodc4.Recordset.Fields!mojodiamademasrafmeghdar = rs1.Fields!mojodiamademasrafmeghdar
        
        rs2.Open "SELECT Sum(meghdar) As rsnumber FROM masrafestandardmavad2 WHERE (idmade='" + Trim(Str(rs1.Fields!idmade)) + "')", db1
          If IsNull(rs2.Fields!rsnumber) = False Then
            Adodc4.Recordset.Fields!masrafteydoremeghdar = rs2.Fields!rsnumber
          Else
            Adodc4.Recordset.Fields!masrafteydoremeghdar = 0
          End If
        rs2.Close
        
        Adodc4.Recordset.Fields!zayeatmeghdar = rs1.Fields!zayeatmeghdar
        Adodc4.Recordset.Fields!mojodipayandoremeghdar = rs1.Fields!mojodipayandoremeghdar
        Adodc4.Recordset.Fields!mojodipayandoremablagh = Val(Adodc4.Recordset.Fields!mojodiamademasrafmeghdar) - (Val(Adodc4.Recordset.Fields!mojodipayandoremeghdar) + Val(Adodc4.Recordset.Fields!zayeatmeghdar))
        
        Adodc4.Recordset.Fields!zayeatmablagh = Val(Adodc4.Recordset.Fields!mojodipayandoremablagh) - Val(Adodc4.Recordset.Fields!masrafteydoremeghdar)
        
        Form4.Adodc1.Recordset.Find "idmavad=" + Trim(Str(rs1.Fields!idmade)), , adSearchForward, 1
        Adodc4.Recordset.Fields!Name = Form4.Adodc1.Recordset.Fields!mavad
        
        Adodc4.Recordset.Update
      
      rs1.MoveNext
    Loop Until rs1.EOF = True
  rs1.Close
db1.Close
Form53.Show
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from kontrolgardeshmes ORDER BY rad"
Adodc1.Refresh

Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "SELECT * FROM Ghardeshmes1 ORDER BY rad"
Adodc3.Refresh

db1.Open Form3.Text10.Text

ProgressBar1.Value = 0
ProgressBar1.Min = 0
ProgressBar1.Max = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
Do
  DoEvents
  ProgressBar1.Value = ProgressBar1.Value + 1
  Select Case Adodc1.Recordset.Fields!rad
    Case 1
      rs1.Open "SELECT * FROM rad WHERE (rad=99998)", db1
        Adodc1.Recordset.Fields!zayeat = rs1.Fields!naghlbebadmeghdar
        Adodc1.Recordset.Update
      rs1.Close
      rs1.Open "SELECT * FROM rad WHERE (rad=99999)", db1
        Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
        Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
        Adodc1.Recordset.Fields!enteghalbade = Val(rs1.Fields!naghlbebadmeghdar) - Val(Adodc1.Recordset.Fields!zayeat)
        Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
        Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin1(Adodc1.Recordset.Fields!rad)

    Case 2
      rs1.Open "SELECT * FROM sanaveye WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin2(Adodc1.Recordset.Fields!rad)
      
    Case 3
      rs1.Open "SELECT * FROM nahaee WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin3(Adodc1.Recordset.Fields!rad)

    Case 4
      rs1.Open "SELECT * FROM Koreh WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin4(Adodc1.Recordset.Fields!rad)

    Case 5
      rs1.Open "SELECT * FROM Taab WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin5(Adodc1.Recordset.Fields!rad)

    Case 6
      rs1.Open "SELECT * FROM Sterander1_6 WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin6(Adodc1.Recordset.Fields!rad)

    Case 7
      rs1.Open "SELECT * FROM Sterander1_36 WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin7(Adodc1.Recordset.Fields!rad)
      
    Case 8
      rs1.Open "SELECT * FROM Sterander1_4 WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin8(Adodc1.Recordset.Fields!rad)

    Case 9
      rs1.Open "SELECT * FROM DramToester WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin9(Adodc1.Recordset.Fields!rad)
      
    Case 10
      rs1.Open "SELECT * FROM Mokhaberat WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin10(Adodc1.Recordset.Fields!rad)
      
    Case 11
      rs1.Open "SELECT Sum(mojodiavalmeghdar) AS mojodiavalmeghdar1, Sum(tolidteydoremeghdar) AS tolidteydoremeghdar1, Sum(naghlbebadmeghdar) AS naghlbebadmeghdar1, Sum(mojodiendmeghdar) AS mojodiendmeghdar1 FROM Exteroder WHERE (nomes='„”  Ê·Ìœ ‘œÂ')", db1
        Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar1
        Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar1
        Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar1
        Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar1
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin11(Adodc1.Recordset.Fields!rad)
      
    Case 12
      rs1.Open "SELECT Sum(mojodiavalmeghdar) AS mojodiavalmeghdar1, Sum(tolidteydoremeghdar) AS tolidteydoremeghdar1, Sum(naghlbebadmeghdar) AS naghlbebadmeghdar1, Sum(mojodiendmeghdar) AS mojodiendmeghdar1 FROM Bastebandi WHERE (nomes='„”  Ê·Ìœ ‘œÂ')", db1
        Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar1
        Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar1
        Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar1
        Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar1
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin12(Adodc1.Recordset.Fields!rad)

    Case 13
      rs1.Open "SELECT * FROM AnbarMahsol WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin13(Adodc1.Recordset.Fields!rad)

    Case 14
      rs1.Open "SELECT * FROM Bancher WHERE (rad=99999)", db1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin14(Adodc1.Recordset.Fields!rad)

    Case 15
      rs1.Open "SELECT Sum(mojodiavalmeghdar) AS mojodiavalmeghdar1, Sum(tolidteydoremeghdar) AS tolidteydoremeghdar1, Sum(naghlbebadmeghdar) AS naghlbebadmeghdar1, Sum(mojodiendmeghdar) AS mojodiendmeghdar1 FROM Exteroder WHERE (nomes='„”  «»ÌœÂ ‘œÂ')", db1
        Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar1
        Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar1
        Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar1
        Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar1
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin15(Adodc1.Recordset.Fields!rad)

    Case 16
      rs1.Open "SELECT Sum(mojodiavalmeghdar) AS mojodiavalmeghdar1, Sum(tolidteydoremeghdar) AS tolidteydoremeghdar1, Sum(naghlbebadmeghdar) AS naghlbebadmeghdar1, Sum(mojodiendmeghdar) AS mojodiendmeghdar1 FROM Bastebandi WHERE (nomes='„”  «»ÌœÂ ‘œÂ')", db1
        Adodc1.Recordset.Fields!mojodiavalmeghdar = rs1.Fields!mojodiavalmeghdar1
        Adodc1.Recordset.Fields!varedeteydoremeghdar = rs1.Fields!tolidteydoremeghdar1
        Adodc1.Recordset.Fields!enteghalbade = rs1.Fields!naghlbebadmeghdar1
        Adodc1.Recordset.Fields!mojodienddore = rs1.Fields!mojodiendmeghdar1
      Adodc1.Recordset.Update
      rs1.Close
      Call KontrolGardeshmes.amin16(Adodc1.Recordset.Fields!rad)

  End Select
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True


Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "SELECT sum(mojodiavalmeghdar) as mojodiavalmeghdar1,sum(varedeteydoremeghdar) as varedeteydoremeghdar1,sum(enteghalbade) as enteghalbade1,sum(zayeat) as zayeat1,sum(mojodienddore)as mojodienddore1, sum(sanaveye1) as sanaveye11,sum(sanaveye2) as sanaveye21,sum(nahaee1) as nahaee11,sum(nahaee2) as nahaee21,sum(Koreh1) as Koreh11,sum(Koreh2) as Koreh21,sum(Taab1) as Taab11,sum(Taab2) as Taab21,sum(Sterander1_61) as Sterander1_611,sum(Sterander1_62) as Sterander1_621,sum(Sterander1_361) as Sterander1_3611,sum(Sterander1_362) as Sterander1_3621,sum(Sterander1_41) as Sterander1_411,sum(Sterander1_42) as Sterander1_421,sum(DramToester1) as DramToester11,sum(DramToester2) as DramToester21,sum(Mokhaberat1) as Mokhaberat11,sum(Mokhaberat2) as Mokhaberat21,sum(Exteroder1) as Exteroder11,sum(Exteroder2) as Exteroder21,sum(Bastebandi1) as Bastebandi11,sum(Bastebandi2) as Bastebandi21,sum(AnbarMahsol1) as AnbarMahsol11,sum(AnbarMahsol2) as AnbarMahsol21 FROM kontrolgardeshmes WHERE (rad <> 999)"
Adodc2.Refresh

Adodc1.Recordset.Find "rad=999", , adSearchForward, 1

Adodc1.Recordset.Fields!mojodiavalmeghdar = Adodc2.Recordset.Fields!mojodiavalmeghdar1
Adodc1.Recordset.Fields!varedeteydoremeghdar = Adodc2.Recordset.Fields!varedeteydoremeghdar1
Adodc1.Recordset.Fields!enteghalbade = Adodc2.Recordset.Fields!enteghalbade1
Adodc1.Recordset.Fields!zayeat = Adodc2.Recordset.Fields!zayeat1
Adodc1.Recordset.Fields!mojodienddore = Adodc2.Recordset.Fields!mojodienddore1
Adodc1.Recordset.Fields!sanaveye1 = Adodc2.Recordset.Fields!sanaveye11
Adodc1.Recordset.Fields!sanaveye2 = Adodc2.Recordset.Fields!sanaveye21
Adodc1.Recordset.Fields!nahaee1 = Adodc2.Recordset.Fields!nahaee11
Adodc1.Recordset.Fields!nahaee2 = Adodc2.Recordset.Fields!nahaee21
Adodc1.Recordset.Fields!Koreh1 = Adodc2.Recordset.Fields!Koreh11
Adodc1.Recordset.Fields!Koreh2 = Adodc2.Recordset.Fields!Koreh21
Adodc1.Recordset.Fields!Taab1 = Adodc2.Recordset.Fields!Taab11
Adodc1.Recordset.Fields!Taab2 = Adodc2.Recordset.Fields!Taab21
Adodc1.Recordset.Fields!Sterander1_61 = Adodc2.Recordset.Fields!Sterander1_611
Adodc1.Recordset.Fields!Sterander1_62 = Adodc2.Recordset.Fields!Sterander1_621
Adodc1.Recordset.Fields!Sterander1_361 = Adodc2.Recordset.Fields!Sterander1_3611
Adodc1.Recordset.Fields!Sterander1_362 = Adodc2.Recordset.Fields!Sterander1_3621
Adodc1.Recordset.Fields!Sterander1_41 = Adodc2.Recordset.Fields!Sterander1_411
Adodc1.Recordset.Fields!Sterander1_42 = Adodc2.Recordset.Fields!Sterander1_421
Adodc1.Recordset.Fields!DramToester1 = Adodc2.Recordset.Fields!DramToester11
Adodc1.Recordset.Fields!DramToester2 = Adodc2.Recordset.Fields!DramToester21
Adodc1.Recordset.Fields!Mokhaberat1 = Adodc2.Recordset.Fields!Mokhaberat11
Adodc1.Recordset.Fields!Mokhaberat2 = Adodc2.Recordset.Fields!Mokhaberat21
Adodc1.Recordset.Fields!Exteroder1 = Adodc2.Recordset.Fields!Exteroder11
Adodc1.Recordset.Fields!Exteroder2 = Adodc2.Recordset.Fields!Exteroder21
Adodc1.Recordset.Fields!Bastebandi1 = Adodc2.Recordset.Fields!Bastebandi11
Adodc1.Recordset.Fields!Bastebandi2 = Adodc2.Recordset.Fields!Bastebandi21
Adodc1.Recordset.Fields!AnbarMahsol1 = Adodc2.Recordset.Fields!AnbarMahsol11
Adodc1.Recordset.Fields!AnbarMahsol2 = Adodc2.Recordset.Fields!AnbarMahsol21
Adodc1.Recordset.Update


Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "SELECT sum(bancher1) as bancher11,sum(bancher2) as bancher21,sum(sum1) as sum11,sum(sum2) as sum21 FROM kontrolgardeshmes WHERE (rad <> 999)"
Adodc2.Refresh

Adodc1.Recordset.Find "rad=999", , adSearchForward, 1

Adodc1.Recordset.Fields!bancher1 = Adodc2.Recordset.Fields!bancher11
Adodc1.Recordset.Fields!bancher2 = Adodc2.Recordset.Fields!bancher21
Adodc1.Recordset.Fields!sum1 = Adodc2.Recordset.Fields!sum11
Adodc1.Recordset.Fields!sum2 = Adodc2.Recordset.Fields!sum21
Adodc1.Recordset.Update

'Adodc1.ConnectionString = Form3.Text10.Text
'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select * from kontrolgardeshmes ORDER BY rad"
'Adodc1.Refresh
DataGrid2.Refresh

'Adodc1.Refresh
DataGrid2.Refresh

Adodc3.Recordset.MoveFirst
Do
  Select Case Adodc3.Recordset.Fields!rad
    Case 1
      rs1.Open "SELECT * FROM kontrolgardeshmes WHERE rad=999", db1
        Adodc3.Recordset.Fields!meghdar = rs1.Fields!mojodiavalmeghdar
        q1(1) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
    
    Case 2
      rs1.Open "SELECT * FROM rad WHERE rad=99999", db1
        Adodc3.Recordset.Fields!meghdar = rs1.Fields!tolidteydoremeghdar
        q1(2) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 3
      rs1.Open "SELECT Sum(vaznmes) As rsnumber FROM Exteroder WHERE (nomes='„”  «»ÌœÂ ‘œÂ')", db1
        Adodc3.Recordset.Fields!meghdar = rs1.Fields!rsnumber
        q1(3) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 4
      Adodc3.Recordset.Fields!meghdar = Val(q1(1)) + Val(q1(2)) + Val(q1(3))
      q1(4) = Adodc3.Recordset.Fields!meghdar
      Adodc3.Recordset.Update
      
    Case 5
      rs1.Open "SELECT * FROM kontrolgardeshmes WHERE rad=999", db1
        Adodc3.Recordset.Fields!meghdar = rs1.Fields!mojodienddore
        q1(5) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 6
      rs1.Open "SELECT * FROM kontrolgardeshmes WHERE rad=999", db1
        Adodc3.Recordset.Fields!meghdar = rs1.Fields!zayeat
        q1(6) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 8
      rs1.Open "SELECT Sum(meghdar) As rsnumber FROM masrafestandardmavad2 WHERE idmade= '1'", db1
        Adodc3.Recordset.Fields!meghdar = Round(rs1.Fields!rsnumber)
        q1(8) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 9
      rs1.Open "SELECT Sum(meghdar) As rsnumber FROM masrafestandardmavad2 WHERE idmade= '30'", db1
        Adodc3.Recordset.Fields!meghdar = Round(rs1.Fields!rsnumber)
        q1(9) = Adodc3.Recordset.Fields!meghdar
        Adodc3.Recordset.Update
      rs1.Close
      
    Case 10
      Adodc3.Recordset.Fields!meghdar = Val(q1(5)) + Val(q1(6)) + Val(q1(8)) + Val(q1(9))
      q1(10) = Adodc3.Recordset.Fields!meghdar
      Adodc3.Recordset.Update
      
    Case 11
      Adodc3.Recordset.Fields!meghdar = Val(q1(4)) - Val(Val(q1(5)) + Val(q1(6)) + Val(q1(8)) + Val(q1(9)))
      q1(11) = Adodc3.Recordset.Fields!meghdar
      Adodc3.Recordset.Update
      
  End Select
  Adodc3.Recordset.MoveNext
Loop Until Adodc3.Recordset.EOF = True

Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "SELECT * FROM Ghardeshmes1 ORDER BY rad"
Adodc3.Refresh
DataGrid3.Refresh

Adodc3.Recordset.Sort = "rad"
Adodc3.Refresh
DataGrid3.Refresh

'Form4.Adodc1.ConnectionString = Form3.Text10.Text
'Form4.Adodc1.CommandType = adCmdUnknown
'Form4.Adodc1.RecordSource = "select * from infomavad where (nosim='1') ORDER BY idmavad"
'Form4.Adodc1.Refresh
'Form4.Adodc1.Recordset.MoveFirst
'Do
'  rs3.Open "Select * From ghardeshmavad Where (nomade=1) and (idmade=" + Trim(Str(Form4.Adodc1.Recordset.Fields!idmavad)) + ")", db1
'  Form4.Adodc1.Recordset.Fields!mvage = rs3.Fields!masrafteydoremeghdar
'  rs3.Close
'
'  rs2.Open "SELECT SUM(meghdar) as amin12 FROM masrafestandardmavad2 WHERE (idmade='" + Trim(Str(Form4.Adodc1.Recordset.Fields!idmavad)) + "')", db1
'  Form4.Adodc1.Recordset.Fields!mesta = rs2.Fields!amin12
'
'  MsgBox rs2.Fields!amin12
'  rs2.Close
'
'  Form4.Adodc1.Recordset.Fields!enheraf = Val(Form4.Adodc1.Recordset.Fields!mvage) - Val(Form4.Adodc1.Recordset.Fields!mesta)
'  Form4.Adodc1.Recordset.Update
'
'  Form4.Adodc1.Recordset.MoveNext
'Loop Until Form4.Adodc1.Recordset.EOF = True
'
'
'Adodc3.Recordset.Find "rad=1", , adSearchForward, 1
'Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Adodc1.Recordset.Fields!mojodiavalmeghdar
'q(1) = Adodc1.Recordset.Fields!mojodiavalmeghdar
'
'Adodc3.Recordset.Find "rad=2", , adSearchForward, 1
'Form9.Adodc3.Recordset.Find "rad=99999", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Form9.Adodc3.Recordset.Fields!tolidteydoremeghdar
'q(2) = Form9.Adodc3.Recordset.Fields!tolidteydoremeghdar
'
'Adodc3.Recordset.Find "rad=3", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Val(q(1)) + Val(q(2))
'q(3) = Val(q(1)) + Val(q(2))
'
'Adodc3.Recordset.Find "rad=4", , adSearchForward, 1
'Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Adodc1.Recordset.Fields!mojodienddore
'q(4) = Adodc1.Recordset.Fields!mojodienddore
'
'Adodc3.Recordset.Find "rad=5", , adSearchForward, 1
'Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Adodc1.Recordset.Fields!zayeat
'q(5) = Adodc1.Recordset.Fields!zayeat
'
'Adodc3.Recordset.Find "rad=6", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Val(q(3)) - Val(q(4)) - Val(q(5))
'q(6) = Val(q(3)) - Val(q(4)) - Val(q(5))
'
'Adodc3.Recordset.Find "rad=7", , adSearchForward, 1
'Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Adodc1.Recordset.Fields!zayeat
'q(7) = Adodc1.Recordset.Fields!zayeat
'
'Adodc3.Recordset.Find "rad=8", , adSearchForward, 1
'Adodc3.Recordset.Fields!meghdar = Val(q(6)) - Val(q(7))
'q(8) = Val(q(6)) - Val(q(7))
'
'Adodc3.Recordset.Update
'
db1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

