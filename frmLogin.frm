VERSION 5.00
Object = "{622AB48B-DF4B-4D9C-AF3A-C94CFE00024D}#2.6#0"; "Adtextbox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login System ...."
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   3625
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4575
         _Version        =   65536
         _ExtentX        =   8070
         _ExtentY        =   1085
         _StockProps     =   15
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdOK 
            Height          =   375
            Left            =   2280
            TabIndex        =   2
            Top             =   120
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Login"
         End
         Begin Threed.SSCommand cmdBatal 
            Height          =   375
            Left            =   3360
            TabIndex        =   5
            Top             =   120
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "&Cancel"
         End
      End
      Begin AdvancedTextBox.adTextBox txtUser 
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
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
      Begin AdvancedTextBox.adTextBox txtPassword 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
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
         PropPasswordChar=   "*"
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "User Name :"
         Height          =   275
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password :"
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   240
         Picture         =   "frmLogin.frx":0442
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
End
End Sub

Private Sub cmdBatal_GotFocus()
cmdOK.Default = False
End Sub

Private Sub cmdOK_Click()

If txtUser.Text = "" Then
    MsgBox "User Name Masih Kosong", vbInformation, "USER NAME"
    txtUser.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If
If txtPassword.Text = "" Then
    MsgBox "PASSWORD Masih Kosong", vbInformation, "PASSWORD"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

If UCase(txtUser.Text) = "EUIS" And UCase(txtPassword.Text) = "EUIS" Then
    frmMenuUtama.Show
    frmMenuUtama.StatusBar1.Panels(1).Text = Trim(txtUser.Text)
    Unload Me
Else
    MsgBox "Maaf, Password Salah. Silahkan cek lagi", vbInformation, "INVALID"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
End If

End Sub

Private Sub Form_Activate()
txtUser.Text = "EUIS"
txtPassword.Text = "EUIS"
cmdOK_Click
End Sub

Private Sub Form_Load()

Dim Atas As Long
Dim Kiri As Long
Atas = (Screen.Height - Me.Height) / 2
Kiri = (Screen.Width - Me.Width) / 2
Me.Move Kiri, Atas

End Sub

Private Sub txtPassword_GotFocus()
cmdOK.Default = True
End Sub

Private Sub txtUser_GotFocus()
cmdOK.Default = False
End Sub

