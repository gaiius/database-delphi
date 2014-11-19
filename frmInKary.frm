VERSION 5.00
Object = "{622AB48B-DF4B-4D9C-AF3A-C94CFE00024D}#2.6#0"; "Adtextbox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInKary 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9825
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   7905
   Begin Threed.SSPanel Panel 
      Height          =   9615
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   16960
      _StockProps     =   15
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin TabDlg.SSTab SSTab1 
         Height          =   4815
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   8493
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "KARYAWAN"
         TabPicture(0)   =   "frmInKary.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "KELUARGA"
         TabPicture(1)   =   "frmInKary.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSFrame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "JAMINAN"
         TabPicture(2)   =   "frmInKary.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSFrame2"
         Tab(2).ControlCount=   1
         Begin Threed.SSFrame SSFrame1 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   38
            Top             =   600
            Width           =   6975
            _Version        =   65536
            _ExtentX        =   12303
            _ExtentY        =   6376
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin VB.ComboBox cbo_jk_kel 
               Height          =   315
               Left            =   1680
               TabIndex        =   22
               Top             =   960
               Width           =   720
            End
            Begin VB.ComboBox cbo_hub 
               Height          =   315
               Left            =   1680
               TabIndex        =   20
               Top             =   240
               Width           =   1800
            End
            Begin TrueOleDBGrid70.TDBGrid Grid2 
               Height          =   2055
               Left            =   120
               TabIndex        =   39
               Top             =   1440
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   3625
               _LayoutType     =   0
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).Locked=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   7
               MultipleLines   =   2
               CellTipsWidth   =   0
               TransparentRowPictures=   -1  'True
               DeadAreaBackColor=   14215660
               RowDividerColor =   16711680
               RowSubDividerColor=   14215660
               DirectionAfterEnter=   0
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=Trebuchet MS"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC08000&,.bold=0"
               _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HFF00FF&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HC080FF&"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=39,.bgcolor=&HC08000&,.bold=0"
               _StyleDefs(23)  =   ":id=11,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(24)  =   ":id=11,.fontname=Trebuchet MS"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000018&"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HC08000&,.bold=-1,.fontsize=975"
               _StyleDefs(29)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(30)  =   ":id=14,.fontname=Trebuchet MS"
               _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(32)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(33)  =   ":id=15,.fontname=Trebuchet MS"
               _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&HFF0000&"
               _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(50)  =   "Named:id=33:Normal"
               _StyleDefs(51)  =   ":id=33,.parent=0"
               _StyleDefs(52)  =   "Named:id=34:Heading"
               _StyleDefs(53)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HFF0000&"
               _StyleDefs(54)  =   ":id=34,.fgcolor=&HFFFFFF&,.wraptext=-1,.locked=-1,.fgpicPosition=0,.appearance=0"
               _StyleDefs(55)  =   ":id=34,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(56)  =   ":id=34,.fontname=Trebuchet MS"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
            Begin AdvancedTextBox.adTextBox txt_nama_kel 
               Height          =   285
               Left            =   1680
               TabIndex        =   21
               Top             =   600
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   503
               CheckKey        =   0
               CheckCase       =   0
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   0   'False
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   0   'False
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin Threed.SSCommand cmdSimpan_Kel 
               Height          =   375
               Left            =   5040
               TabIndex        =   23
               Top             =   960
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Simpan"
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   375
               Left            =   6000
               TabIndex        =   43
               Top             =   960
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Hapus"
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Jenis Kelamin :"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nama : "
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hubungan : "
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   1455
            End
         End
         Begin Threed.SSFrame Frame1 
            Height          =   4035
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   6975
            _Version        =   65536
            _ExtentX        =   12303
            _ExtentY        =   7117
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin VB.ComboBox cbo_kawin 
               Height          =   315
               Left            =   1800
               TabIndex        =   7
               Top             =   2520
               Width           =   675
            End
            Begin VB.ComboBox cbo_jk 
               Height          =   315
               Left            =   1800
               TabIndex        =   3
               Top             =   1080
               Width           =   720
            End
            Begin VB.ComboBox cbo_divisi 
               Height          =   315
               Left            =   1800
               TabIndex        =   8
               Top             =   2880
               Width           =   945
            End
            Begin VB.ComboBox cbo_jab 
               Height          =   315
               Left            =   1785
               TabIndex        =   9
               Top             =   3240
               Width           =   945
            End
            Begin MSComCtl2.DTPicker tgl_lhr 
               Height          =   330
               Left            =   1800
               TabIndex        =   4
               Top             =   1440
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   582
               _Version        =   393216
               Format          =   59703297
               CurrentDate     =   39995
            End
            Begin AdvancedTextBox.adTextBox txt_nik 
               Height          =   285
               Left            =   1800
               TabIndex        =   1
               Top             =   360
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   503
               CheckKey        =   0
               CheckCase       =   1
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   0   'False
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   -1  'True
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin AdvancedTextBox.adTextBox txt_nama 
               Height          =   285
               Left            =   1800
               TabIndex        =   2
               Top             =   720
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   503
               CheckKey        =   0
               CheckCase       =   0
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   0   'False
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   0   'False
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin AdvancedTextBox.adTextBox txt_alamat 
               Height          =   285
               Left            =   1800
               TabIndex        =   5
               Top             =   1800
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   503
               CheckKey        =   0
               CheckCase       =   0
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   0   'False
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   0   'False
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin AdvancedTextBox.adTextBox txt_kota 
               Height          =   285
               Left            =   1800
               TabIndex        =   6
               Top             =   2160
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   503
               CheckKey        =   0
               CheckCase       =   0
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   0   'False
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   0   'False
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin VB.Label lbl_Gapok 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000080FF&
               Height          =   255
               Left            =   1770
               TabIndex        =   53
               Top             =   3600
               Width           =   2040
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Gaji Pokok :"
               Height          =   255
               Left            =   240
               TabIndex        =   52
               Top             =   3600
               Width           =   1455
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Status Kawin :"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   2520
               Width           =   1455
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Kota :"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   2160
               Width           =   1455
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Alamat :"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   1800
               Width           =   1455
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tanggal Lahir :"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Jenis Kelamin :"
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nama :"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.I.K :"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Divisi :"
               Height          =   255
               Left            =   240
               TabIndex        =   30
               Top             =   2880
               Width           =   1455
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Jabatan :"
               Height          =   255
               Left            =   225
               TabIndex        =   29
               Top             =   3240
               Width           =   1455
            End
            Begin VB.Label lbl_divisi 
               BackColor       =   &H000080FF&
               Height          =   255
               Left            =   2835
               TabIndex        =   28
               Top             =   2880
               Width           =   3480
            End
            Begin VB.Label lbl_kawin 
               BackColor       =   &H000080FF&
               Height          =   255
               Left            =   2835
               TabIndex        =   27
               Top             =   2565
               Width           =   3480
            End
            Begin VB.Label lbl_jab 
               BackColor       =   &H000080FF&
               Height          =   255
               Left            =   2835
               TabIndex        =   26
               Top             =   3240
               Width           =   3480
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   44
            Top             =   600
            Width           =   6975
            _Version        =   65536
            _ExtentX        =   12303
            _ExtentY        =   6376
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin VB.ComboBox cbo_th 
               Height          =   315
               Left            =   1680
               TabIndex        =   45
               Top             =   240
               Width           =   1080
            End
            Begin TrueOleDBGrid70.TDBGrid Grid3 
               Height          =   2535
               Left            =   120
               TabIndex        =   46
               Top             =   960
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   4471
               _LayoutType     =   0
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).Locked=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
               PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   7
               MultipleLines   =   2
               CellTipsWidth   =   0
               TransparentRowPictures=   -1  'True
               DeadAreaBackColor=   14215660
               RowDividerColor =   16711680
               RowSubDividerColor=   14215660
               DirectionAfterEnter=   0
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=Trebuchet MS"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC08000&,.bold=0"
               _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HFF00FF&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HC080FF&"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=39,.bgcolor=&HC08000&,.bold=0"
               _StyleDefs(23)  =   ":id=11,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(24)  =   ":id=11,.fontname=Trebuchet MS"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000018&"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HC08000&,.bold=-1,.fontsize=975"
               _StyleDefs(29)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(30)  =   ":id=14,.fontname=Trebuchet MS"
               _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=900,.italic=0"
               _StyleDefs(32)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(33)  =   ":id=15,.fontname=Trebuchet MS"
               _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&HFF0000&"
               _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(50)  =   "Named:id=33:Normal"
               _StyleDefs(51)  =   ":id=33,.parent=0"
               _StyleDefs(52)  =   "Named:id=34:Heading"
               _StyleDefs(53)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HFF0000&"
               _StyleDefs(54)  =   ":id=34,.fgcolor=&HFFFFFF&,.wraptext=-1,.locked=-1,.fgpicPosition=0,.appearance=0"
               _StyleDefs(55)  =   ":id=34,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(56)  =   ":id=34,.fontname=Trebuchet MS"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
            Begin AdvancedTextBox.adTextBox txt_biaya 
               Height          =   285
               Left            =   1680
               TabIndex        =   47
               Top             =   600
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   503
               CheckKey        =   4
               CheckCase       =   0
               CheckValidation =   0
               PropText        =   ""
               PropFriendlyName=   ""
               PropAlignment   =   0
               PropAppearance  =   1
               BeginProperty PropFontname {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropBorderStyle =   1
               PropForeColor   =   -2147483640
               PropBackColor   =   -2147483643
               PropDiameterKey =   0
               PropDateSeparator=   0
               PropMinLength   =   0
               SwitchEnabled   =   -1  'True
               SwitchLocked    =   -1  'True
               SwitchRequired  =   0   'False
               SwitchSelectionFocus=   0   'False
               HiddenFontSize  =   8,25
               HiddenFontBold  =   0   'False
               HiddenFontItalic=   0   'False
               HiddenFontStrikethru=   0   'False
               HiddenFontUnderline=   0   'False
               PropMaxLength   =   0
               PropPasswordChar=   ""
               PropCustomCharacterString=   ""
               SwitchAutoSkip  =   0   'False
               PropAdditionalBackKey=   0
               PropAdditionalNextKey=   13
               SwitchAllowMinus=   0   'False
               SwitchAllowThousandSeparator=   0   'False
               SwitchAllowDecimalSeparator=   0   'False
               PropErrorMessage=   ""
               PropRegularExpression=   ""
            End
            Begin Threed.SSCommand cmdSimpanJamin 
               Height          =   375
               Left            =   5040
               TabIndex        =   48
               Top             =   360
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Simpan"
            End
            Begin Threed.SSCommand SSCommand3 
               Height          =   375
               Left            =   6000
               TabIndex        =   49
               Top             =   360
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Hapus"
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tahun Anggaran :"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Jaminan :"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   600
               Width           =   1455
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   240
         TabIndex        =   12
         Top             =   5580
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   1508
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   2
         Begin Threed.SSCommand cmdTambah 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "T&ambah"
         End
         Begin Threed.SSCommand cmdBatal 
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Batal"
         End
         Begin Threed.SSCommand cmdSimpan 
            Height          =   375
            Left            =   2760
            TabIndex        =   10
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Simpan"
         End
         Begin Threed.SSCommand cmdUbah 
            Height          =   375
            Left            =   3720
            TabIndex        =   14
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Ubah"
         End
         Begin Threed.SSCommand cmdHapus 
            Height          =   375
            Left            =   4680
            TabIndex        =   15
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Hapus"
         End
         Begin Threed.SSCommand cmdTutup 
            Height          =   375
            Left            =   5640
            TabIndex        =   16
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Tutup"
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3045
         Left            =   240
         TabIndex        =   17
         Top             =   6495
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   5371
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid70.TDBGrid Grid1 
            Height          =   2895
            Left            =   120
            TabIndex        =   18
            Top             =   75
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5106
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   7
            MultipleLines   =   2
            CellTipsWidth   =   0
            TransparentRowPictures=   -1  'True
            DeadAreaBackColor=   14215660
            RowDividerColor =   16711680
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   0
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Trebuchet MS"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC08000&,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HFF00FF&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HC080FF&"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=39,.bgcolor=&HC08000&,.bold=0"
            _StyleDefs(23)  =   ":id=11,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(24)  =   ":id=11,.fontname=Trebuchet MS"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000018&"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HC08000&,.bold=-1,.fontsize=975"
            _StyleDefs(29)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(30)  =   ":id=14,.fontname=Trebuchet MS"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(32)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=15,.fontname=Trebuchet MS"
            _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&HFF0000&"
            _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Named:id=33:Normal"
            _StyleDefs(51)  =   ":id=33,.parent=0"
            _StyleDefs(52)  =   "Named:id=34:Heading"
            _StyleDefs(53)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HFF0000&"
            _StyleDefs(54)  =   ":id=34,.fgcolor=&HFFFFFF&,.wraptext=-1,.locked=-1,.fgpicPosition=0,.appearance=0"
            _StyleDefs(55)  =   ":id=34,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(56)  =   ":id=34,.fontname=Trebuchet MS"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "INPUT DATA KARYAWAN "
         ForeColor       =   12632064
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         Alignment       =   4
      End
   End
End
Attribute VB_Name = "frmInKary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset
Public rsKawin As ADODB.Recordset
Public rsDivisi As ADODB.Recordset
Public rsJab As ADODB.Recordset
Public rsKel As ADODB.Recordset
Public rsJamin As ADODB.Recordset

Private Sub cbo_divisi_Click()
CariDivisi
End Sub

Private Sub cbo_divisi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    CariDivisi
End If
KeyAscii = 0
End Sub

Private Sub cbo_hub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    CariJab
End If
KeyAscii = 0
End Sub

Private Sub cbo_jab_Click()
CariJab
End Sub

Private Sub cbo_jab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    CariJab
End If
KeyAscii = 0
End Sub

Private Sub cbo_jk_kel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    CariJab
End If
KeyAscii = 0
End Sub

Private Sub cbo_jk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
KeyAscii = 0
End Sub

Private Sub cbo_kawin_Click()
CariKawin
End Sub

Private Sub cbo_kawin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    CariKawin
End If
KeyAscii = 0
End Sub

Private Sub cbo_th_Click()
On Error Resume Next
txt_biaya.Text = Format(Val(Abs(lbl_Gapok.Caption)) * 10, "###,###,###")
End Sub

Private Sub cbo_th_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdBatal_Click()
settombol True
txt_nik.Enabled = True
Kosong
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Err:
If txt_nik.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_divisi.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Divisi", vbInformation, "Info"
    Exit Sub
End If

If cbo_jab.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Jabatan", vbInformation, "Info"
    Exit Sub
End If

Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tb_karyawan WHERE nik='" & Trim(txt_nik.Text) & "'", cn, 1, 2
If rs.EOF Then
    MsgBox "Maaf, Tidak ada data yang dapat dihapus", vbInformation, "Info"
    Exit Sub
Else
    If MsgBox("Anda yakin akan menghapus data : " & txt_nik.Text & " : " & txt_nama.Text, vbYesNo, "Hapus Data") = vbYes Then
        Set rs = New ADODB.Recordset
        rs.Open "DELETE FROM tb_karyawan WHERE nik='" & Trim(txt_nik.Text) & "'", cn, 1, 2
        Kosong
        TampilData
        TampilDataKel
    End If
End If

Exit Sub
Err:
If Err.Number = -2147217873 Then
    MsgBox "Maaf, Data tidak dapat dihapus, karena ada Relasi", vbInformation, "Info"
End If
End Sub

Private Sub cmdSimpan_Click()
If txt_nik.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    txt_nik.SetFocus
    Exit Sub
End If

If cbo_divisi.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Divisi", vbInformation, "Info"
    cbo_divisi.SetFocus
    Exit Sub
End If

If cbo_jab.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Jabatan", vbInformation, "Info"
    cbo_jab.SetFocus
    Exit Sub
End If


Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tb_karyawan WHERE nik='" & Trim(txt_nik.Text) & "'", cn, 1, 2
If rs.EOF Then
    With rs
        .AddNew
        !nik = Trim(txt_nik.Text)
        !nm_karyawan = Trim(txt_nama.Text)
        !jenis_kelamin = cbo_jk.Text
        !tgl_lhr = Format(tgl_lhr.Value, "dd/MM/yyyy")
        !kd_kawin = cbo_kawin.Text
        !alamat = Trim(txt_alamat.Text)
        !kota = Trim(txt_kota.Text)
        !kd_divisi = cbo_divisi.Text
        !kd_jab = cbo_jab.Text
        .Update
    End With
Else
    If MsgBox("Maaf, NIK sudah ada. Apakah Anda akan mengubahnya", vbYesNo, "Ubah Data") = vbYes Then
        With rs
            !nm_karyawan = Trim(txt_nama.Text)
            !jenis_kelamin = cbo_jk.Text
            !tgl_lhr = Format(tgl_lhr.Value, "dd/MM/yyyy")
            !kd_kawin = cbo_kawin.Text
            !alamat = Trim(txt_alamat.Text)
            !kota = Trim(txt_kota.Text)
            !kd_divisi = cbo_divisi.Text
            !kd_jab = cbo_jab.Text
            .Update
        End With
    Else
        txt_nik.SetFocus
    End If
End If
settombol True
TampilData
End Sub

Private Sub cmdSimpan_Kel_Click()
If txt_nik.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_hub.Text = Empty Then
    MsgBox "Maaf, Hubungan keluarga tidak boleh kosong", vbInformation, "Info"
    cbo_hub.SetFocus
    Exit Sub
End If

If txt_nama_kel.Text = Empty Then
    MsgBox "Maaf, Nama Keluarga tidak boleh kosong", vbInformation, "Info"
    txt_nama_kel.SetFocus
    Exit Sub
End If

If cbo_kawin = "K1" Then
    MsgBox "Maaf, Status Perkawinan belum Kawin", vbInformation, "Info"
    Exit Sub
End If

Set rsKel = New ADODB.Recordset
rsKel.Open "SELECT * FROM tb_keluarga WHERE nik='" & txt_nik.Text & "' AND hub_kel='" & cbo_hub.Text & "'", cn, 1, 2
If rsKel.EOF Then
    rsKel.AddNew
    rsKel!nik = Trim(txt_nik.Text)
    rsKel!hub_kel = Trim(cbo_hub.Text)
    rsKel!nama_kel = Trim(txt_nama_kel.Text)
    rsKel!jenis_kelamin = Trim(cbo_jk_kel.Text)
    rsKel.Update
Else
    If MsgBox("Maaf, Hubungan Keluarga sudah ada, apakah anda akan mengubahnnya?", vbYesNo, "Ubah Data") = vbYes Then
        rsKel!nama_kel = Trim(txt_nama_kel.Text)
        rsKel!jenis_kelamin = Trim(cbo_jk_kel.Text)
        rsKel.Update
    Else
        cbo_hub.SetFocus
    End If
End If

settombol True
TampilDataKel
Kosong_K
End Sub

Private Sub cmdSimpanJamin_Click()
If txt_nik.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_th.Text = Empty Then
    MsgBox "Maaf, Tahun Anggaran tidak boleh kosong", vbInformation, "Info"
    cbo_th.SetFocus
    Exit Sub
End If

Set rsJamin = New ADODB.Recordset
rsJamin.Open "SELECT * FROM tb_jaminan WHERE nik='" & Trim(txt_nik.Text) & "' AND th_anggaran='" & Trim(cbo_th.Text) & "'", cn, 1, 2
If rsJamin.EOF Then
    With rsJamin
        .AddNew
        !nik = Trim(txt_nik.Text)
        !th_anggaran = Trim(cbo_th.Text)
        !biaya_jaminan = Abs(txt_biaya.Text)
        .Update
    End With
Else
    If MsgBox("Biaya jaminan sudah ada, apakah anda akan mengubahnya?", vbYesNo, "Ubah Data") = vbYes Then
        With rsJamin
            !biaya_jaminan = Abs(txt_biaya.Text)
            .Update
        End With
    Else
        cbo_th.SetFocus
    End If
End If

TampilData
TampilJamin
End Sub

Private Sub cmdTambah_Click()
settombol False
Kosong
txt_nik.Enabled = True
txt_nik.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub cmdUbah_Click()
If txt_nik.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_divisi.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Divisi", vbInformation, "Info"
    Exit Sub
End If

If cbo_jab.Text = Empty Then
    MsgBox "Maaf, Anda belum memilih Jabatan", vbInformation, "Info"
    Exit Sub
End If

settombol False
txt_nik.Enabled = False
txt_nama.SetFocus
End Sub

Private Sub Form_Load()
Ketengah Me

settombol True
TampilData
ItemCombo
tgl_lhr.Value = Format(Now, "dd/MM/yyyy")

cbo_jk.AddItem "L"
cbo_jk.AddItem "P"

cbo_hub.AddItem "Suami"
cbo_hub.AddItem "Istri"
cbo_hub.AddItem "Anak Ke-1"
cbo_hub.AddItem "Anak Ke-2"

cbo_jk_kel.AddItem "L"
cbo_jk_kel.AddItem "P"

For i = 2008 To Year(Now)
    cbo_th.AddItem (Str(i))
Next i

End Sub

Sub ItemCombo()
Set rsKawin = New ADODB.Recordset
rsKawin.Open "SELECT * FROM tb_status_kawin ORDER BY kd_kawin", cn, 1, 2
If rsKawin.EOF Then
    cbo_kawin.Text = Empty
Else
    rsKawin.MoveFirst
    Do Until rsKawin.EOF
        cbo_kawin.AddItem rsKawin!kd_kawin
        rsKawin.MoveNext
    Loop
End If

Set rsDivisi = New ADODB.Recordset
rsDivisi.Open "SELECT * FROM tb_divisi ORDER BY kd_divisi", cn, 1, 2
If rsDivisi.EOF Then
    cbo_divisi.Text = Empty
Else
    rsDivisi.MoveFirst
    Do Until rsDivisi.EOF
        cbo_divisi.AddItem rsDivisi!kd_divisi
        rsDivisi.MoveNext
    Loop
End If

Set rsJab = New ADODB.Recordset
rsJab.Open "SELECT * FROM tb_jab ORDER BY kd_jab", cn, 1, 2
If rsJab.EOF Then
    cbo_jab.Text = Empty
Else
    rsJab.MoveFirst
    Do Until rsJab.EOF
        cbo_jab.AddItem rsJab!kd_jab
        rsJab.MoveNext
    Loop
End If
End Sub

Sub CariKawin()
Set rsKawin = New ADODB.Recordset
rsKawin.Open "SELECT * FROM tb_status_kawin WHERE kd_kawin='" & Trim(cbo_kawin.Text) & "'", cn, 1, 2
If rsKawin.EOF Then
    lbl_kawin.Caption = Empty
Else
    lbl_kawin.Caption = rsKawin!status_kawin
End If
End Sub

Sub CariDivisi()
Set rsDivisi = New ADODB.Recordset
rsDivisi.Open "SELECT * FROM tb_divisi WHERE kd_divisi='" & Trim(cbo_divisi.Text) & "'", cn, 1, 2
If rsDivisi.EOF Then
    lbl_divisi.Caption = Empty
Else
    lbl_divisi.Caption = rsDivisi!divisi
End If
End Sub

Sub CariJab()
Set rsJab = New ADODB.Recordset
rsJab.Open "SELECT * FROM tb_jab WHERE kd_jab='" & Trim(cbo_jab.Text) & "'", cn, 1, 2
If rsJab.EOF Then
    lbl_jab.Caption = Empty
    lbl_Gapok.Caption = Empty
Else
    lbl_jab.Caption = rsJab!jabatan
    lbl_Gapok.Caption = Format(rsJab!gapok, "###,###,###")
End If
End Sub

Sub settombol(bval As Boolean)
cmdTambah.Enabled = bval
cmdBatal.Enabled = Not bval
cmdSimpan.Enabled = Not bval
cmdUbah.Enabled = bval
cmdHapus.Enabled = bval
cmdTutup.Enabled = bval

Frame1.Enabled = Not bval

End Sub

Sub Kosong()
tgl_lhr.Value = Format(Now, "dd/MM/yyyy")
txt_nik.Text = Empty
txt_nama.Text = Empty
cbo_jk.Text = Empty
txt_alamat.Text = Empty
txt_kota.Text = Empty
cbo_kawin.Text = Empty: lbl_kawin.Caption = Empty
cbo_divisi.Text = Empty: lbl_divisi.Caption = Empty
cbo_jab.Text = Empty: lbl_jab.Caption = Empty

TampilDataKel
End Sub

Sub Kosong_K()
cbo_hub.Text = Empty
txt_nama_kel.Text = Empty
cbo_jk_kel.Text = Empty
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IsNull(Grid1.Columns(0).Value) Then
      GetList Grid1.Columns(0).Value
    End If
End Sub

Private Function GetList(pID As String)
On Error Resume Next
      
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tb_karyawan WHERE nik='" & pID & "'", cn, 1, 2
    If rs.EOF Then
        GetList = False
    Else
    With rs
        txt_nik.Text = ValidNull(!nik)
        txt_nama.Text = ValidNull(!nm_karyawan)
        cbo_jk.Text = ValidNull(!jenis_kelamin)
        tgl_lhr.Value = Format(!tgl_lhr, "dd/MM/yyyy")
        txt_alamat.Text = ValidNull(!alamat)
        txt_kota.Text = ValidNull(!kota)
        cbo_kawin.Text = ValidNull(!kd_kawin)
        cbo_divisi.Text = ValidNull(!kd_divisi)
        cbo_jab.Text = ValidNull(!kd_jab)
        
        CariKawin
        CariDivisi
        CariJab
        
        TampilDataKel
        TampilJamin
        GetList = True
        
    End With
    End If
End Function


Sub TampilData()
Set rs = New ADODB.Recordset
rs.Open "SELECT nik,nm_karyawan,jenis_kelamin,kd_jab FROM tb_karyawan", cn, 1, 2

Set Grid1.DataSource = rs
DoRefreshGrid Grid1, rs
With Grid1
    .Columns(0).Caption = "NIK"
    .Columns(0).Width = 1000
    .Columns(1).Caption = "Nama Pasien"
    .Columns(1).Width = 3500
    .Columns(2).Caption = "Sex"
    .Columns(2).Width = 500
    .Columns(3).Caption = "Jabatan"
    .Columns(3).Width = 1500
End With
End Sub

Sub TampilDataKel()
Set rsKel = New ADODB.Recordset
rsKel.Open "SELECT nik,hub_kel,Nama_kel,jenis_kelamin FROM tb_keluarga WHERE nik='" & Trim(txt_nik.Text) & "'", cn, 1, 2

Set Grid2.DataSource = rsKel

With Grid2
    .Columns(0).Visible = False
    .Columns(1).Caption = "Hubungan"
    .Columns(1).Width = 1500
    .Columns(2).Caption = "Nama Keluarga"
    .Columns(2).Width = 3500
    .Columns(3).Caption = "Sex"
    .Columns(3).Width = 500
End With
End Sub

Sub TampilJamin()
Set rsJamin = New ADODB.Recordset
rsJamin.Open "SELECT nik,th_anggaran,biaya_jaminan FROM tb_jaminan WHERE nik='" & Trim(txt_nik.Text) & "'", cn, 1, 2

Set Grid3.DataSource = rsJamin

With Grid3
    .Columns(0).Visible = False
    .Columns(1).Caption = "Tahun"
    .Columns(1).Width = 1500
    .Columns(2).Caption = "Jaminan"
    .Columns(2).Width = 3500
    .Columns(2).NumberFormat = "###,###,###"
End With
End Sub

Private Sub tgl_lhr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub
