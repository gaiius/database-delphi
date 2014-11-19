VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLapSPJPK 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6435
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
   ScaleHeight     =   3780
   ScaleWidth      =   6435
   Begin Threed.SSPanel SSPanel1 
      Height          =   3615
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6376
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
      BevelInner      =   1
      Begin Threed.SSFrame SSFrame1 
         Height          =   2895
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   5106
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
         Begin VB.ComboBox cbo_jab 
            Height          =   315
            Left            =   2040
            TabIndex        =   10
            Top             =   1185
            Width           =   2175
         End
         Begin VB.ComboBox cbo_Divisi 
            Height          =   315
            Left            =   2040
            TabIndex        =   9
            Top             =   870
            Width           =   2175
         End
         Begin VB.ComboBox cbo_Kawin 
            Height          =   315
            Left            =   2040
            TabIndex        =   8
            Top             =   555
            Width           =   1230
         End
         Begin VB.ComboBox cbo_NIK 
            Height          =   315
            Left            =   2040
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Semuannya"
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   1500
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Per Jabatan"
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   1185
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Per Divisi"
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   870
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Per Tahun"
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   555
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Per NIK"
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   855
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   5535
            _Version        =   65536
            _ExtentX        =   9763
            _ExtentY        =   1508
            _StockProps     =   15
            BackColor       =   16777152
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
            Begin Threed.SSCommand cmdCetak 
               Height          =   375
               Left            =   3480
               TabIndex        =   12
               Top             =   240
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Cetak"
            End
            Begin Threed.SSCommand cmdTutup 
               Height          =   375
               Left            =   4440
               TabIndex        =   13
               Top             =   240
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   661
               _StockProps     =   78
               Caption         =   "&Tutup"
            End
         End
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4680
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "LAPORAN SPJPK "
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
Attribute VB_Name = "frmLapSPJPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset
Public rsKary As ADODB.Recordset
Public rsKawin As ADODB.Recordset
Public rsDivisi As ADODB.Recordset
Public rsJab As ADODB.Recordset

Private Sub cbo_divisi_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_jab_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_kawin_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_NIK_KeyDown(KeyCode As Integer, Shift As Integer)
KeyAscii = 0
End Sub

Private Sub cmdCetak_Click()

If Option1.Value = True Then
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM VIEW_LAP_SPJPK WHERE NIK='" & Trim(cbo_NIK.Text) & "'", cn, 1, 2
    If rs.EOF Then
        MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
        Exit Sub
    Else
    With CrystalReport1
        .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
        .ReportFileName = App.Path & "\Lap_SPJPK.rpt"
        .WindowState = crptMaximized
        .SelectionFormula = "{VIEW_LAP_SPJPK.NIK}='" & Trim(cbo_NIK.Text) & "'"
        .RetrieveDataFiles
        .Action = 1
        .Reset
    End With
    End If
ElseIf Option2.Value = True Then
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM VIEW_LAP_SPJPK WHERE year(tgl_spjpk)= " & Trim(cbo_Kawin.Text) & " ", cn, 1, 2
    If rs.EOF Then
        MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
        Exit Sub
    Else
    With CrystalReport1
        .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
        .ReportFileName = App.Path & "\Lap_SPJPK.rpt"
        .WindowState = crptMaximized
        .SelectionFormula = "year({VIEW_LAP_SPJPK.tgl_spjpk})= " & Trim(cbo_Kawin.Text) & " "
        .RetrieveDataFiles
        .Action = 1
        .Reset
    End With
    End If
ElseIf Option3.Value = True Then
    If cbo_Divisi.Text = "" Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM LAP_SPJPK_DIVISI ", cn, 1, 2
        If rs.EOF Then
            MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
            Exit Sub
        Else
        With CrystalReport1
            .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
            .ReportFileName = App.Path & "\Lap_SPJPK_Divisi.rpt"
            .WindowState = crptMaximized
            .SelectionFormula = ""
            .RetrieveDataFiles
            .Action = 1
            .Reset
        End With
        End If
    Else
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM LAP_SPJPK_DIVISI WHERE divisi='" & Trim(cbo_Divisi.Text) & "'", cn, 1, 2
        If rs.EOF Then
            MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
            Exit Sub
        Else
        With CrystalReport1
            .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
            .ReportFileName = App.Path & "\Lap_SPJPK_Divisi.rpt"
            .WindowState = crptMaximized
            .SelectionFormula = "{LAP_SPJPK_DIVISI.divisi}='" & Trim(cbo_Divisi.Text) & "'"
            .RetrieveDataFiles
            .Action = 1
            .Reset
        End With
        End If
    End If
ElseIf Option4.Value = True Then
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM view_Jaminan WHERE jabatan='" & Trim(cbo_jab.Text) & "'", cn, 1, 2
    If rs.EOF Then
        MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
        Exit Sub
    Else
    With CrystalReport1
        .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
        .ReportFileName = App.Path & "\Lap_SPJPK.rpt"
        .WindowState = crptMaximized
        .SelectionFormula = "{VIEW_LAP_SPJPK.jabatan}='" & Trim(cbo_jab.Text) & "'"
        .RetrieveDataFiles
        .Action = 1
        .Reset
    End With
    End If
ElseIf Option5.Value = True Then
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM view_karyawan", cn, 1, 2
    If rs.EOF Then
        MsgBox "Maaf, tidak ada data yang diCetak", vbInformation, "Info"
        Exit Sub
    Else
    With CrystalReport1
        .LogOnServer "p2ssql.dll", "KRAMOTAX", "db_euis_ta", "sa", "sa"
        .ReportFileName = App.Path & "\Lap_SPJPK.rpt"
        .WindowState = crptMaximized
        .SelectionFormula = ""
        .RetrieveDataFiles
        .Action = 1
        .Reset
    End With
    End If
End If
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Ketengah Me

Option5.Value = True

SetInput False, False, False, False

ItemCombo
End Sub

Sub ItemCombo()
Set rsKary = New ADODB.Recordset
rsKary.Open "SELECT * FROM view_karyawan ORDER BY NIK", cn, 1, 2
If rsKary.EOF Then
    cbo_NIK.Text = Empty
Else
    rsKary.MoveFirst
    Do Until rsKary.EOF
        cbo_NIK.AddItem rsKary!nik
        rsKary.MoveNext
    Loop
End If

For i = 2008 To Year(Now)
    cbo_Kawin.AddItem (Str(i))
Next i

'Set rsKawin = New ADODB.Recordset
'rsKawin.Open "SELECT * FROM tb_status_kawin ORDER BY kd_kawin", cn, 1, 2
'If rsKawin.EOF Then
'    cbo_Kawin.Text = Empty
'Else
'    While Not rsKawin.EOF
'        cbo_Kawin.AddItem rsKawin!status_kawin
'        rsKawin.MoveNext
'    Wend
'End If

Set rsDivisi = New ADODB.Recordset
rsDivisi.Open "SELECT * FROM tb_divisi ORDER BY kd_divisi", cn, 1, 2
If rsDivisi.EOF Then
    cbo_Divisi.Text = Empty
Else
    While Not rsDivisi.EOF
        cbo_Divisi.AddItem rsDivisi!divisi
        rsDivisi.MoveNext
    Wend
End If

Set rsJab = New ADODB.Recordset
rsJab.Open "SELECT * FROM tb_jab ORDER BY kd_jab", cn, 1, 2
If rsJab.EOF Then
    cbo_jab.Text = Empty
Else
    While Not rsJab.EOF
        cbo_jab.AddItem rsJab!jabatan
        rsJab.MoveNext
    Wend
End If
End Sub

Sub SetInput(bval1 As Boolean, bval2 As Boolean, bval3 As Boolean, bval4 As Boolean)
cbo_NIK.Enabled = bval1
cbo_Kawin.Enabled = bval2
cbo_Divisi.Enabled = bval3
cbo_jab.Enabled = bval4
End Sub

Private Sub Option1_Click()
SetInput True, False, False, False
End Sub

Private Sub Option2_Click()
SetInput False, True, False, False
End Sub

Private Sub Option3_Click()
SetInput False, False, True, False
End Sub

Private Sub Option4_Click()
SetInput False, False, False, True
End Sub

Private Sub Option5_Click()
SetInput False, False, False, False
End Sub


