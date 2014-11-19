VERSION 5.00
Object = "{622AB48B-DF4B-4D9C-AF3A-C94CFE00024D}#2.6#0"; "Adtextbox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInSPJPK 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7815
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   7815
   Begin Threed.SSPanel Panel 
      Height          =   8625
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   15214
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   7335
         _Version        =   65536
         _ExtentX        =   12938
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
            Left            =   360
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
            Left            =   1335
            TabIndex        =   9
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
            Left            =   2310
            TabIndex        =   6
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
            Left            =   3285
            TabIndex        =   10
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
            Left            =   4260
            TabIndex        =   11
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
            Left            =   6210
            TabIndex        =   12
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Tutup"
         End
         Begin Threed.SSCommand cmdCetak 
            Height          =   375
            Left            =   5235
            TabIndex        =   35
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Cetak"
         End
      End
      Begin Threed.SSFrame Frame1 
         Height          =   3900
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   7335
         _Version        =   65536
         _ExtentX        =   12938
         _ExtentY        =   6879
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
         Begin VB.ComboBox txt_praktek 
            Height          =   315
            Left            =   1785
            TabIndex        =   4
            Top             =   3195
            Width           =   3660
         End
         Begin VB.ComboBox cbo_Pasien 
            Height          =   315
            Left            =   1800
            TabIndex        =   2
            Top             =   2520
            Width           =   1635
         End
         Begin VB.ComboBox cbo_NIK 
            Height          =   315
            Left            =   1800
            TabIndex        =   1
            Top             =   1080
            Width           =   1635
         End
         Begin MSComCtl2.DTPicker tgl_SPJPK 
            Height          =   330
            Left            =   1800
            TabIndex        =   34
            Top             =   720
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            _Version        =   393216
            Format          =   59768833
            CurrentDate     =   39995
         End
         Begin AdvancedTextBox.adTextBox txt_dokter 
            Height          =   285
            Left            =   1785
            TabIndex        =   5
            Top             =   3510
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
         Begin AdvancedTextBox.adTextBox txt_tujuan 
            Height          =   285
            Left            =   1785
            TabIndex        =   3
            Top             =   2880
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tujuan :"
            Height          =   255
            Left            =   225
            TabIndex        =   38
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nama Dokter :"
            Height          =   255
            Left            =   225
            TabIndex        =   37
            Top             =   3510
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Praktek :"
            Height          =   255
            Left            =   225
            TabIndex        =   36
            Top             =   3195
            Width           =   1455
         End
         Begin VB.Label lbl_nama_pasien 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   3555
            TabIndex        =   33
            Top             =   2565
            Width           =   3480
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pasien :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lbl_S_Jaminan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   5280
            TabIndex        =   31
            Top             =   2160
            Width           =   1800
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sisa Jaminan :"
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lbl_T_Jaminan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   29
            Top             =   2160
            Width           =   1800
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Jaminan :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lbl_Jaminan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   5280
            TabIndex        =   27
            Top             =   1800
            Width           =   1800
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Jaminan :"
            Height          =   255
            Left            =   3720
            TabIndex        =   26
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lbl_Gapok 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   25
            Top             =   1800
            Width           =   1800
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Gaji Pokok :"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lbl_No 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label lbl_Divisi 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   22
            Top             =   1440
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Divisi / Jabatan :"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lbl_Nama 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   3555
            TabIndex        =   20
            Top             =   1125
            Width           =   3480
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "NIK :"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "No.SPJPK"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tanggal :"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3045
         Left            =   120
         TabIndex        =   16
         Top             =   5475
         Width           =   7335
         _Version        =   65536
         _ExtentX        =   12938
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
            TabIndex        =   17
            Top             =   75
            Width           =   7095
            _ExtentX        =   12515
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
            MultipleLines   =   0
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
         TabIndex        =   18
         Top             =   0
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "INPUT DATA SPJPK "
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
Attribute VB_Name = "frmInSPJPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset
Public rsKary As ADODB.Recordset
Public rsKel As ADODB.Recordset
Public rsJamin As ADODB.Recordset
Public rsTJamin As ADODB.Recordset

Private Sub cbo_NIK_Click()
CariKaryawan
End Sub

Private Sub cbo_NIK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    CariKaryawan
End If
KeyAscii = 0
End Sub

Private Sub cbo_Pasien_Click()
CariPasien
End Sub

Private Sub cbo_Pasien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    CariPasien
End If
KeyAscii = 0
End Sub

Private Sub cmdBatal_Click()
settombol True
cbo_NIK.Enabled = True
Kosong
End Sub

Private Sub cmdHapus_Click()
On Error GoTo Err
If cbo_NIK.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If lbl_Nama.Caption = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_Pasien.Text = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If lbl_nama_pasien.Caption = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tb_spjpk WHERE no_spjpk='" & lbl_No.Caption & "'", cn, 1, 2
If rs.EOF Then
    MsgBox "Maaf, Tidak ada data yang dapat dihapus", vbInformation, "Info"
    Exit Sub
Else
    If MsgBox("Anda yakin akan menghapus No.SPJPK : " & lbl_No.Caption, vbYesNo, "Hapus Data") = vbYes Then
        Set rs = New ADODB.Recordset
        rs.Open "DELETE FROM tb_spjpk WHERE no_spjpk='" & lbl_No.Caption & "'", cn, 1, 2
        Kosong
        TampilData
        settombol True
    End If
End If
Exit Sub
Err:
If Err.Number = -2147217873 Then
    MsgBox "Maaf, Data tidak dapat dihapus, karena ada Relasi", vbInformation, "Info"
End If
End Sub

Private Sub cmdSimpan_Click()
If cbo_NIK.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    cbo_NIK.SetFocus
    Exit Sub
End If

If lbl_Nama.Caption = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    cbo_NIK.SetFocus
    Exit Sub
End If

If cbo_Pasien.Text = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    cbo_Pasien.SetFocus
    Exit Sub
End If

If lbl_nama_pasien.Caption = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    cbo_Pasien.SetFocus
    Exit Sub
End If

If lbl_S_Jaminan.Caption <= 0 Then
    MsgBox "Maaf, Anda tidak dapat mengajukan SPJPK, Lihat sisa Jaminan", vbInformation, "Info"
    Exit Sub
End If

If txt_praktek.Text = Empty Then
    MsgBox "Maaf, Tempat Praktek Tidak boleh Kosong", vbInformation, "Info"
    txt_praktek.SetFocus
    Exit Sub
End If

If txt_dokter.Text = Empty Then
    MsgBox "Maaf, Dokter Tidak boleh Kosong", vbInformation, "Info"
    txt_dokter.SetFocus
    Exit Sub
End If

Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tb_spjpk WHERE no_spjpk='" & lbl_No.Caption & "'", cn, 1, 2
If rs.EOF Then
    With rs
        .AddNew
        !no_spjpk = Trim(lbl_No.Caption)
        !tgl_SPJPK = Format(tgl_SPJPK.Value, "dd/MM/yyyy")
        !nik = Trim(cbo_NIK.Text)
        !hub_kel = Trim(cbo_Pasien.Text)
        !nama_pasien = Trim(lbl_nama_pasien.Caption)
        !praktek = Trim(txt_praktek.Text)
        !dokter = Trim(txt_dokter.Text)
        !tujuan = Trim(txt_tujuan.Text)
        .Update
    End With
Else
    If MsgBox("Nomor SPJPK sudah ada apakah Anda akan mengubahnnya?", vbYesNo, "Ubah Data") = vbYes Then
        With rs
            !hub_kel = Trim(cbo_Pasien.Text)
            !nama_pasien = Trim(lbl_nama_pasien.Caption)
            !praktek = Trim(txt_praktek.Text)
            !dokter = Trim(txt_dokter.Text)
            !tujuan = Trim(txt_tujuan.Text)
            .Update
        End With
    End If
End If

settombol True
TampilData
End Sub

Private Sub cmdTambah_Click()
settombol False
Kosong
NoSPJPK
cbo_NIK.Enabled = True
cbo_NIK.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub cmdUbah_Click()
If cbo_NIK.Text = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If lbl_Nama.Caption = Empty Then
    MsgBox "Maaf, NIK tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If cbo_Pasien.Text = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

If lbl_nama_pasien.Caption = Empty Then
    MsgBox "Maaf, Nama Pasien tidak boleh kosong", vbInformation, "Info"
    Exit Sub
End If

settombol False
cbo_NIK.Enabled = False
cbo_Pasien.SetFocus
End Sub

Private Sub Form_Load()
Ketengah Me

settombol True
TampilData
ItemCombo
End Sub


Sub ItemCombo()
txt_praktek.AddItem "Sps. Kebidanan"
txt_praktek.AddItem "Sps. Anak"
txt_praktek.AddItem "Sps. Gigi"
txt_praktek.AddItem "Sps. Radiologi"
txt_praktek.AddItem "Sps. Pskiater"
txt_praktek.AddItem "Sps. Kulit"

Set rsKary = New ADODB.Recordset
rsKary.Open "SELECT * FROM tb_karyawan ORDER BY nik", cn, 1, 2
If rsKary.EOF Then
    cbo_NIK.Text = Empty
Else
    rsKary.MoveFirst
    Do Until rsKary.EOF
        cbo_NIK.AddItem rsKary!nik
        rsKary.MoveNext
    Loop
End If

cbo_Pasien.AddItem "Karyawan"
cbo_Pasien.AddItem "Suami"
cbo_Pasien.AddItem "Istri"
cbo_Pasien.AddItem "Anak Ke-1"
cbo_Pasien.AddItem "Anak Ke-2"
'Set rsKel = New ADODB.Recordset
'rsKel.Open "SELECT * FROM tb_keluarga ", cn, 1, 2
'If rsKel.EOF Then
'    cbo_Pasien.Text = Empty
'Else
'    cbo_Pasien.AddItem "Karyawan"
'    While Not rsKel.EOF
'        cbo_Pasien.AddItem rsKel!hub_kel
'        rsKel.MoveNext
'    Wend
'End If
End Sub

Sub CariKaryawan()
Set rsKary = New ADODB.Recordset
rsKary.Open "SELECT * FROM view_karyawan WHERE nik='" & Trim(cbo_NIK.Text) & "'", cn, 1, 2
If rsKary.EOF Then
    lbl_Nama.Caption = Empty
    lbl_divisi.Caption = Empty
Else
    With rsKary
        lbl_Nama.Caption = !nm_karyawan
        lbl_divisi.Caption = ValidNull(!divisi) & " / " & ValidNull(!jabatan)
        lbl_Gapok.Caption = Format(!gapok, "###,###,###")
    End With
    
    Set rsJamin = New ADODB.Recordset
    rsJamin.Open "SELECT * FROM tb_jaminan WHERE nik='" & Trim(cbo_NIK.Text) & "' AND th_anggaran='" & Year(Now) & "'", cn, 1, 2
    If rsJamin.EOF Then
        lbl_Jaminan.Caption = 0
    Else
        lbl_Jaminan.Caption = Format(rsJamin!biaya_jaminan, "###,###,###")
    End If
    
    Set rsTJamin = New ADODB.Recordset
    rsTJamin.Open "SELECT * FROM view_total_rincian WHERE nik='" & Trim(cbo_NIK.Text) & "' AND th='" & Year(Now) & "'", cn, 1, 2
    If rsTJamin.EOF Then
        lbl_T_Jaminan.Caption = 0
        lbl_S_Jaminan.Caption = Format(Val(Abs(lbl_Jaminan.Caption)) - Val(Abs(lbl_T_Jaminan.Caption)), "###,###,###")
    Else
        lbl_T_Jaminan.Caption = Format(rsTJamin!jml, "###,###,###")
        lbl_S_Jaminan.Caption = Format(Val(Abs(lbl_Jaminan.Caption)) - Val(Abs(lbl_T_Jaminan.Caption)), "###,###,###")
    End If
End If
End Sub

Sub CariPasien()
If cbo_Pasien.Text = "Karyawan" Then
    lbl_nama_pasien.Caption = lbl_Nama.Caption
Else
    Set rsKel = New ADODB.Recordset
    rsKel.Open "SELECT * FROM tb_keluarga WHERE nik='" & Trim(cbo_NIK.Text) & "' AND hub_kel='" & cbo_Pasien.Text & "'", cn, 1, 2
    If rsKel.EOF Then
        lbl_nama_pasien.Caption = ""
    Else
        lbl_nama_pasien.Caption = rsKel!nama_kel
    End If
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
tgl_SPJPK.Value = Format(Now, "dd/MM/yyyy")
cbo_NIK.Text = Empty
lbl_Nama.Caption = Empty
lbl_divisi.Caption = Empty
lbl_Gapok.Caption = Empty
lbl_Jaminan.Caption = Empty
lbl_T_Jaminan.Caption = Empty: lbl_S_Jaminan.Caption = Empty
cbo_Pasien.Text = Empty: lbl_nama_pasien.Caption = Empty
txt_praktek.Text = Empty
txt_dokter.Text = Empty
txt_tujuan.Text = Empty
End Sub

Sub NoSPJPK()
Dim txtRes As String
Dim NomorKwitansi As String
Dim Tit As String

Tit = Format(Now, "dd") & Format(Now, "MM") & Format(Now, "yy")
Set rs = New ADODB.Recordset
rs.Open "select * from tb_spjpk order by no_spjpk", cn, 1, adLockOptimistic

If Not rs.EOF Then
rs.MoveLast
NomorKwitansi = Val(Right(rs("no_spjpk"), 6)) + 1
Select Case Abs(NomorKwitansi)
    Case 0 To 9
        lbl_No.Caption = Tit & "00000" & NomorKwitansi
    Case 10 To 99
        lbl_No.Caption = Tit & "0000" & NomorKwitansi
    Case 100 To 999
        lbl_No.Caption = Tit & "000" & NomorKwitansi
    Case 1000 To 9999
        lbl_No.Caption = Tit & "00" & NomorKwitansi
    Case 10000 To 99999
        lbl_No.Caption = Tit & "0" & NomorKwitansi
    Case 10000 To 999999
        lbl_No.Caption = Tit & NomorKwitansi
End Select
Else
lbl_No.Caption = Tit & "000001"
End If

End Sub

Sub TampilData()
Set rs = New ADODB.Recordset
rs.Open "SELECT no_spjpk,tgl_spjpk,nik,nm_karyawan,hub_kel,nama_pasien,tujuan,dokter FROM view_spjpk", cn, 1, 2

Set Grid1.DataSource = rs
DoRefreshGrid Grid1, rs
With Grid1
    .Columns(0).Caption = "No.SPJPK"
    .Columns(0).Width = 1500
    .Columns(1).Caption = "Tanggal"
    .Columns(1).Width = 1200
    .Columns(2).Caption = "NIK"
    .Columns(2).Width = 1000
    .Columns(3).Caption = "Nama Karyawan"
    .Columns(3).Width = 2500
    .Columns(4).Caption = "Hub Kel"
    .Columns(4).Width = 1500
    .Columns(5).Caption = "Nama Pasien"
    .Columns(5).Width = 2500
End With
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IsNull(Grid1.Columns(0).Value) Then
      GetList Grid1.Columns(0).Value
    End If
End Sub

Private Function GetList(pID As String)
On Error Resume Next
  Dim rs As New ADODB.Recordset
  
    Set rs = New ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM view_spjpk WHERE no_spjpk='" & pID & "'", cn, 1, 2
    If rs.EOF Then
        GetList = False
    Else
    With rs
        lbl_No.Caption = Trim(!no_spjpk)
        tgl_SPJPK.Value = Format(!tgl_SPJPK, "dd/MM/yyyy")
        cbo_NIK.Text = ValidNull(!nik)
        cbo_Pasien.Text = ValidNull(!hub_kel)
        lbl_nama_pasien.Caption = ValidNull(!nama_pasien)
        txt_praktek.Text = ValidNull(!praktek)
        txt_dokter.Text = ValidNull(!dokter)
        txt_tujuan.Text = ValidNull(!tujuan)
        
        CariKaryawan
        CariPasien
                
        GetList = True
    End With
    End If
End Function

Private Sub txt_praktek_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
KeyAscii = 0
End Sub
