VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CODEJO~3.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmBarang 
   BorderStyle     =   0  'None
   Caption         =   $"FrmBarang.frx":0000
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoProduk 
      Height          =   330
      Left            =   8280
      Top             =   6120
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin BasKomponen.BasForm BasForm1 
      Height          =   5640
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9948
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   ":: Barang ::"
      Object.ToolTipText     =   ":: Barang ::"
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   3615
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   6376
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Barang"
         Columns(0).DataField=   "KodeBarang"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Produk"
         Columns(1).DataField=   "NamaProduk"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Barang"
         Columns(2).DataField=   "NamaBarang"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Satuan"
         Columns(3).DataField=   "Satuan"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Harga Beli"
         Columns(4).DataField=   "HargaBeli"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Harga Jual"
         Columns(5).DataField=   "HargaJual"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Stock"
         Columns(6).DataField=   "Stock"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   529
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   8421376
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         Appearance      =   2
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         GroupByCaption  =   "Keterangan"
         DeadAreaBackColor=   14215660
         RowDividerColor =   8454143
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H7DBBFF&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.borderColor=&H80000013&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFF8080&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000005&,.fgcolor=&H0&,.bold=0"
         _StyleDefs(26)  =   ":id=13,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=13,.fontname=Verdana"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=37,.bgcolor=&H555555&"
         _StyleDefs(30)  =   ":id=14,.fgcolor=&H37D7FF&,.bold=-1,.fontsize=600,.italic=0,.underline=0"
         _StyleDefs(31)  =   ":id=14,.strikethrough=0,.charset=255"
         _StyleDefs(32)  =   ":id=14,.fontname=Terminal"
         _StyleDefs(33)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.fgcolor=&HFFFF&,.borderColor=&H80FF80&"
         _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(37)  =   ":id=17,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=17,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.fgcolor=&HFFFF&"
         _StyleDefs(40)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF&"
         _StyleDefs(41)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=37,.bgcolor=&H80FFFF&"
         _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   ":id=35,.wraptext=0,.locked=0"
         _StyleDefs(80)  =   "Named:id=36:Selected"
         _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   ":id=36,.borderColor=&H80000013&"
         _StyleDefs(83)  =   "Named:id=37:Caption"
         _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(85)  =   "Named:id=38:HighlightRow"
         _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HA00000&,.borderColor=&H800040&"
         _StyleDefs(87)  =   "Named:id=39:EvenRow"
         _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(89)  =   "Named:id=40:OddRow"
         _StyleDefs(90)  =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(91)  =   "Named:id=41:RecordSelector"
         _StyleDefs(92)  =   ":id=41,.parent=34"
         _StyleDefs(93)  =   "Named:id=42:FilterBar"
         _StyleDefs(94)  =   ":id=42,.parent=33,.bgcolor=&H80FFFF&,.fgcolor=&H0&"
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   6600
         TabIndex        =   25
         ToolTipText     =   "Cetak barcode"
         Top             =   2520
         Width           =   1215
         _Version        =   851972
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Barcode"
         BackColor       =   8454143
         Appearance      =   6
         BorderGap       =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TxtStock 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
         _Version        =   393216
         _ExtentX        =   2778
         _ExtentY        =   556
         Calculator      =   "FrmBarang.frx":00E0
         Caption         =   "FrmBarang.frx":0100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmBarang.frx":016C
         Keys            =   "FrmBarang.frx":018A
         Spin            =   "FrmBarang.frx":01D4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TxtJual 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
         _Version        =   393216
         _ExtentX        =   2778
         _ExtentY        =   556
         Calculator      =   "FrmBarang.frx":01FC
         Caption         =   "FrmBarang.frx":021C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmBarang.frx":0288
         Keys            =   "FrmBarang.frx":02A6
         Spin            =   "FrmBarang.frx":02F0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TxtBeli 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
         _Version        =   393216
         _ExtentX        =   2778
         _ExtentY        =   556
         Calculator      =   "FrmBarang.frx":0318
         Caption         =   "FrmBarang.frx":0338
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmBarang.frx":03A4
         Keys            =   "FrmBarang.frx":03C2
         Spin            =   "FrmBarang.frx":040C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbSatuan 
         Height          =   345
         Left            =   2160
         TabIndex        =   4
         Tag             =   "Kode"
         Top             =   1920
         Width           =   2535
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         BevelType       =   0
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Row.Count       =   12
         Row(0)          =   "Batang"
         Row(1)          =   "Biji"
         Row(2)          =   "Lusin"
         Row(3)          =   "Botol"
         Row(4)          =   "Bungkus"
         Row(5)          =   "Kardus"
         Row(6)          =   "Lembar"
         Row(7)          =   "Karton"
         Row(8)          =   "Saset"
         Row(9)          =   "Potong"
         Row(10)         =   "Paket"
         Row(11)         =   "Tablet"
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   4419
         Columns(0).Caption=   "Satuan"
         Columns(0).Name =   "Satuan"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbProduk 
         Height          =   345
         Left            =   2160
         TabIndex        =   3
         Tag             =   "Kode"
         Top             =   1560
         Width           =   2535
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbGrup 
         Height          =   345
         Left            =   2160
         TabIndex        =   2
         Tag             =   "Kode"
         Top             =   1200
         Width           =   2535
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   855
         Left            =   360
         ScaleHeight     =   795
         ScaleWidth      =   8595
         TabIndex        =   16
         Top             =   4560
         Width           =   8655
         Begin XtremeSuiteControls.PushButton CmdAdd 
            Height          =   615
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Tambah data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":0434
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdEdit 
            Height          =   615
            Left            =   720
            TabIndex        =   9
            ToolTipText     =   "Edit data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":5DDE
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdDelete 
            Height          =   615
            Left            =   1320
            TabIndex        =   10
            ToolTipText     =   "Hapus data"
            Top             =   80
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":B788
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdSave 
            Height          =   615
            Left            =   6720
            TabIndex        =   11
            ToolTipText     =   "Simpan data"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":11132
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdCancel 
            Height          =   615
            Left            =   7320
            TabIndex        =   12
            ToolTipText     =   "Cancel"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":16ADC
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdQuit 
            Height          =   615
            Left            =   7920
            TabIndex        =   13
            ToolTipText     =   "Exit"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":1C486
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.PushButton CmdPrint 
            Height          =   615
            Left            =   3960
            TabIndex        =   14
            ToolTipText     =   "Cetak data barang"
            Top             =   75
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   1085
            _StockProps     =   79
            BackColor       =   8454143
            Appearance      =   6
            Picture         =   "FrmBarang.frx":1C7A0
            BorderGap       =   0
         End
      End
      Begin VB.TextBox TxtKode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1920
         Width           =   2475
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   4395
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   7
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   6
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   5
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   4
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   3
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grup Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   0
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label LB_Reload 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   6000
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label LB_Reload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   210
         Index           =   1
         Left            =   480
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   840
         Width           =   1065
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   3285
         Index           =   0
         Left            =   360
         Top             =   600
         Width           =   8505
      End
   End
   Begin MSAdodcLib.Adodc AdoGrup 
      Height          =   330
      Left            =   6960
      Top             =   6120
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Dim RsTempAFI As New ADODB.Recordset
Dim Cols As TrueOleDBGrid70.Columns
Sub bersih()
TxtKode = ""
TxtNama = ""
CmbGrup.Text = ""
CmbProduk = ""
CmbSatuan = ""
TxtBeli.Value = 0
TxtJual.Value = 0
TxtStock.Value = 0
End Sub
Function KodeAuto()
SQL = "Select KodeBarang from barang where kodeProduk='" & Trim(CmbProduk.Columns(0).Text) & _
    "' order by kodeProduk Desc"
Set RsFind = DbCon.Execute(SQL)
If RsFind.BOF Then
   KodeAuto = Trim(CmbProduk.Columns(0).Text) + "00001"
Else
   KodeAuto = Trim(CmbProduk.Columns(0).Text) + Format(CInt(Right(RsFind!KodeBarang, 5)) + 1, "00000")
End If
End Function

Private Sub CmbGrup_DropDown()
AdoGrup.RecordSource = ""
SQL = "Select kodeJenis as [Kode Grup],namaJenis as [Nama Jenis] from jenis order by namaJenis"
Set RsFind = DbCon.Execute(SQL)
CmbGrup.Reset
If RsFind.BOF Then Exit Sub
AdoGrup.RecordSource = SQL
AdoGrup.Refresh
With CmbGrup
        .DataSourceList = AdoGrup
        .DataFieldList = "nama jenis"
        .Columns(1).Width = 3000
        .Columns(0).Visible = False
End With
End Sub

Private Sub CmbProduk_Click()
TxtKode = KodeAuto
End Sub

Private Sub CmbProduk_DropDown()
AdoProduk.RecordSource = ""
SQL = "Select KodeProduk as [Kode Produk],NamaProduk as [Nama Produk] from produk where kodeGrup ='" & _
    Trim(CmbGrup.Columns(0).Text) & "'"
Set RsFind = DbCon.Execute(SQL)
CmbProduk.Reset
If RsFind.BOF Then Exit Sub
AdoProduk.RecordSource = SQL
AdoProduk.Refresh
With CmbProduk
        .DataSourceList = AdoProduk
        .DataFieldList = "nama produk"
        .Columns(1).Width = 3000
        .Columns(0).Visible = False
End With
End Sub

Private Sub CmdAdd_Click()
tombol False
Edit = False
TxtNama.SetFocus
End Sub

Private Sub CmdCancel_Click()
tombol True
RefreshData
bersih

CmbGrup.Enabled = True
CmbProduk.Enabled = True
End Sub

Private Sub CmdDelete_Click()
If MsgBox("Yakin akan menghapus data ini?" & vbCrLf & "" _
            & "KODE BARANG: " & Trim(Grid.Columns(0).Text) + vbCrLf & "" _
            & "NAMA BARANG: " & Trim(Grid.Columns(2).Text) + vbCrLf & "", _
         vbYesNo + vbQuestion) = vbYes Then
    SQL = "delete from barang where kodeBarang ='" & Trim(Grid.Columns(0).Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data terhapus"
    RefreshData
End If
End Sub

Private Sub CmdEdit_Click()
Grid_DblClick
End Sub

Private Sub CmdPrint_Click()
Grid.ExportToFile App.Path & _
                "\Daftar barang.xls", False
ShellEx App.Path & "\Daftar barang.xls", Owner:=Me.hwnd
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
If Not checkIsi Then Exit Sub

If Not Edit Then
    SQL = "insert into barang (kodeBarang,kodeProduk,NamaBarang,Satuan,HargaBeli, " & _
        " HargaJual,Stock) values ('" & Trim(TxtKode.Text) & "','" & _
        Trim(CmbProduk.Columns(0).Text) & "' ,'" & Trim(TxtNama) & "','" & Trim(CmbSatuan.Text) & _
        "'," & TxtBeli.Value & "," & TxtJual.Value & "," & TxtStock.Value & ")"
Else
    SQL = "update barang Set namabarang='" & Trim(TxtNama.Text) & "', satuan='" & _
        Trim(CmbSatuan.Text) & "',hargabeli=" & TxtBeli.Value & ", hargajual = " & _
        TxtJual.Value & ", stock=" & TxtStock.Value & " where kodebarang ='" & Trim(TxtKode.Text) & "'"
    CmbGrup.Enabled = True
    CmbProduk.Enabled = True
End If
DbCon.Execute SQL
MsgBox "Data tersimpan."
tombol True
bersih
RefreshData
End Sub

Function checkIsi() As Boolean
If Trim(TxtNama) = "" Then
    MsgBox "Nama Barang masih kosong."
    TxtNama.SetFocus
    checkIsi = False
ElseIf Trim(CmbGrup.Text) = "" Or Not CmbGrup.IsItemInList Then
    MsgBox "Grup belum dipilih"
    CmbGrup.SetFocus
    checkIsi = False
ElseIf Trim(CmbProduk.Text) = "" Or Not CmbProduk.IsItemInList Then
    MsgBox "Produk belum dipilih"
    CmbProduk.SetFocus
    checkIsi = False
ElseIf Trim(CmbSatuan.Text) = "" Or Not CmbSatuan.IsItemInList Then
    MsgBox "Satuan belum dipilih"
    CmbSatuan.SetFocus
    checkIsi = False
ElseIf TxtBeli.Value = 0 Then
    MsgBox "Harga Beli masih 0."
    TxtBeli.SetFocus
    checkIsi = False
ElseIf TxtJual.Value = 0 Then
    MsgBox "Harga Jual masih 0."
    TxtJual.SetFocus
    checkIsi = False
ElseIf TxtStock.Value = 0 Then
    MsgBox "Stock masih 0."
    TxtStock.SetFocus
    checkIsi = False
Else
    checkIsi = True
End If
    
End Function

Private Sub Form_Load()
Me.Height = BasForm1.Height
Me.Width = BasForm1.Width

AdoGrup.ConnectionString = ConDb
AdoProduk.ConnectionString = ConDb

CmbSatuan.ZOrder vbSendToBack
CmbProduk.ZOrder vbSendToBack
CmbGrup.ZOrder vbSendToBack

Grid.Left = 240
bersih
tombol True
RefreshData
End Sub

Sub RefreshData()
Set Grid.DataSource = Nothing
SQL = "SELECT * FROM Barang INNER JOIN Produk ON Barang.KodeProduk = Produk.KodeProduk"

Set RsTempAFI = DbCon.Execute(SQL)
Grid.DataSource = RsTempAFI
Grid.Refresh

Grid.Columns(0).Visible = False
Grid.Columns(1).Alignment = dbgLeft
Grid.Columns(1).Width = 3000
Grid.Columns(2).Alignment = dbgLeft
Grid.Columns(2).Width = 3000
Grid.Columns(3).Alignment = dbgLeft
Grid.Columns(3).Width = 2000
Grid.Columns(4).Alignment = dbgLeft
Grid.Columns(4).Width = 2000
Grid.Columns(5).Alignment = dbgLeft
Grid.Columns(5).Width = 2000
Grid.Columns(6).Alignment = dbgLeft
Grid.Columns(6).Width = 2000
End Sub

Sub tombol(Status As Boolean)
CmdAdd.Visible = Status
CmdEdit.Visible = Status
CmdDelete.Visible = Status
CmdPrint.Visible = Status

CmdSave.Visible = Not Status
CmdCancel.Visible = Not Status

Grid.Visible = Status
End Sub

Private Sub Grid_DblClick()
SQL = "SELECT Barang.KodeBarang,(SELECT NamaJenis From Jenis WHERE (KodeJenis = " & _
    " LEFT(Barang.KodeProduk, 4))) AS Namajenis, Produk.NamaProduk, Barang.NamaBarang, " & _
    " Barang.Satuan, Barang.HargaBeli, Barang.HargaJual , Barang.Stock FROM Barang INNER JOIN " & _
    " Produk ON Barang.KodeProduk = Produk.KodeProduk where kodebarang='" & _
    Trim(Grid.Columns(0).Text) & " '"
Set RsFind = DbCon.Execute(SQL)

TxtKode = RsFind!KodeBarang
TxtNama = RsFind!namabarang
CmbGrup.Text = RsFind!namajenis
CmbProduk.Text = RsFind!NamaProduk
CmbSatuan.Text = RsFind!satuan
TxtBeli.Value = RsFind!Hargabeli
TxtJual.Value = RsFind!hargajual
TxtStock.Value = RsFind!stock

CmbGrup.Enabled = False
CmbProduk.Enabled = False

tombol False
Edit = True
End Sub
Private Sub Grid_FilterChange()

On Error GoTo Vb_Error
   Set Cols = Grid.Columns
   Dim c As Integer
   c = Grid.Col
   'If Cols(7).FilterText <> "" Then Cols(7).FilterText = ""
   Grid.HoldFields
   RsTempAFI.Filter = getFilter()
   If RsTempAFI.EOF Then
      MsgBox "Data yang dimaksud tidak dapat ditemukan !", vbInformation
      For Each Col In Grid.Columns
         Col.FilterText = ""
      Next Col
      RefreshData
      Grid.SetFocus
      Grid.Col = c
      Exit Sub
   End If
   Grid.Col = c
   Grid.EditActive = True
   Exit Sub

Vb_Error:
   If Err.Number = 3001 Then MsgBox "Ada Kesalahan Inputan", vbInformation
   For Each Col In Grid.Columns
      Col.FilterText = ""
   Next Col
   'Resume Next
End Sub

Private Function getFilter() As String
1     On Error GoTo Vb_Error
         Dim tmp As String
         Dim n As Integer
2        For Each Col In Cols
3           If Trim(Col.FilterText) <> "" Then
4              n = n + 1
5              If n > 1 Then
6                 tmp = tmp & " AND "
7              End If
8              If Col.ExternalEditor = "tgltime" Then
9                 If Len(Col.FilterText) <> 10 Then Exit Function
10                tmp = tmp & Col.DataField & " > '" & Col.FilterText & "' and " & Col.DataField & " < '" & CDate(CLng(CDate(Col.FilterText)) + 1) & "'"
11             ElseIf Col.ExternalEditor = "tgl" Then
12                If Len(Col.FilterText) <> 10 Then Exit Function
13                tmp = tmp & Col.DataField & " = '" & Col.FilterText & "'"
14             ElseIf Col.ExternalEditor = "Bit" Then
15                If Len(Col.FilterText) <> 1 Then Exit Function
16                tmp = tmp & Col.DataField & " = " & Col.FilterText & ""
17             Else
18                tmp = tmp & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
19             End If
20          End If
21       Next Col
22       getFilter = tmp
23       Exit Function
Vb_Error:
24       Resume Next
End Function

Private Sub PushButton1_Click()
FrmBarcode.Show 1
End Sub
