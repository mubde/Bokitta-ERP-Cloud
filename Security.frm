VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{81FEA250-2DA5-40F7-A3F1-6F8532B748DB}#1.0#0"; "ciaXPPanel30.ocx"
Object = "{88F7F54F-F24B-4B64-B0E0-2454E1E6DA40}#1.0#0"; "ciaXPButton30.ocx"
Object = "{A7E76481-D275-422D-A506-9F4C890EE7D7}#1.0#0"; "ciaXPFrame30.ocx"
Object = "{268F1C9C-A0AF-4B47-BB71-8D9162E3480C}#1.0#0"; "ciaXPText30.ocx"
Object = "{8ADF6797-780A-487D-BC69-26E5708C9A3F}#1.0#0"; "ciaXPSelection30.ocx"
Object = "{8CD7576A-DE38-4ABA-A9D1-206DB09FED98}#1.0#0"; "ciaXPLabel30.ocx"
Object = "{CB232AFF-530C-47D1-BCBA-9450F6AED806}#1.0#0"; "ciaXPSideBarMenu30.ocx"
Object = "{59FAAC82-8FD1-41B8-8597-8C4A376A2CDC}#1.0#0"; "ciaXPDP30.ocx"
Object = "{61C20119-5677-48E5-9D43-CBF5F7B39FA0}#1.0#0"; "ciaXPCombo30.ocx"
Object = "{85F497B1-6519-447B-A5DB-1FCF65C55151}#1.0#0"; "ciaXPImage30.ocx"
Object = "{D28A06DF-024F-4EE4-A3A3-F4E44CBE6D12}#1.0#0"; "ciaXPMultiLineText30.ocx"
Object = "{B9F109CF-3065-4A6D-BED8-2D64A160EF37}#3.1#0"; "eLifePListL.ocx"
Object = "{2E8C9854-D084-45DD-9E54-A9BE7FF10D96}#1.0#0"; "eLifeTabStripL.ocx"
Begin VB.Form frmSecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Settings"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   Icon            =   "Security.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin ciaXPPanel30.XPPanel30 XPPanel301 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   1720
      BackStyle       =   1
      BackColor       =   14737632
      BorderLeftStyle =   2
      BorderRightStyle=   2
      BorderTopStyle  =   2
      BorderBottomStyle=   2
      LicValid        =   -1  'True
      Begin ciaXPLabel30.XPLabel30 lblComment 
         Height          =   645
         Left            =   60
         Top             =   270
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1138
         BackStyle       =   1
         BackColor       =   16777215
         Border          =   0   'False
         Caption         =   $"Security.frx":038A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         WordWrap        =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPLabel30.XPLabel30 lblTitle 
         Height          =   255
         Left            =   30
         Top             =   30
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   450
         AutoSize        =   -1  'True
         AutoSelectTheme =   -1  'True
         BackStyle       =   1
         BackColor       =   16777215
         Border          =   0   'False
         Caption         =   "0043006F006E0064006900740069006F006E00730020006200790020"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         LicValid        =   -1  'True
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   0
         Picture         =   "Security.frx":056E
         Top             =   0
         Width           =   8235
      End
   End
   Begin ciaXPPanel30.XPPanel30 XPPanel302 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   4785
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   1244
      AutoSelectTheme =   -1  'True
      BackStyle       =   1
      BackColor       =   16764331
      BorderBottomStyle=   0
      LicValid        =   -1  'True
      Begin ciaXPMultiLineText30.XPMultiLineText30 lblInfo 
         Height          =   525
         Left            =   690
         TabIndex        =   40
         Top             =   90
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   926
         BackColor       =   16764331
         BorderWidth     =   0
         ButtonCaption   =   ""
         ButtonPicture   =   "Security.frx":1F560
         ButtonMaskColor =   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   8388608
         LockedBackColor =   16764331
         Locked          =   -1  'True
         LockedBackColor =   16764331
         LockedForeColor =   8388608
         MousePointer    =   0
         MouseIcon       =   "Security.frx":1FF5A
         MultiLine       =   -1  'True
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPImage30.XPImage30 img 
         Height          =   315
         Left            =   3990
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Picture         =   "Security.frx":1FF76
         ImageHeight     =   420
         ImageWidth      =   420
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnInfo 
         Height          =   500
         Left            =   90
         TabIndex        =   38
         Top             =   90
         Visible         =   0   'False
         Width           =   550
         _ExtentX        =   979
         _ExtentY        =   873
         AutoSelectTheme =   -1  'True
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         BackStyle       =   0
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnCancel 
         Height          =   315
         Left            =   6180
         TabIndex        =   2
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         AutoSelectTheme =   -1  'True
         Caption         =   "002600430061006E00630065006C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 cmdOK 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         AutoSelectTheme =   -1  'True
         Caption         =   "0026004F006B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
   End
   Begin eLifeTabStripL.eTabStripL tbsOptions 
      Align           =   1  'Align Top
      Height          =   3810
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   6720
      Tabs            =   9
      TabCaption0     =   "00540061006200200030"
      TabEnabled0     =   -1  'True
      TabColor0       =   12640511
      TabBorderColor0 =   14737632
      Tab(0)Controls  =   4
      Tab(0)Ctrl(1)ID =   "XPFrame301"
      Tab(0)Ctrl(1)Name=   "XPFrame301"
      Tab(0)Ctrl(2)ID =   "cboPolicy(0)"
      Tab(0)Ctrl(2)Index=   0
      Tab(0)Ctrl(2)Name=   "cboPolicy"
      Tab(0)Ctrl(3)ID =   "fraPGPOptions"
      Tab(0)Ctrl(3)Name=   "fraPGPOptions"
      Tab(0)Ctrl(4)ID =   "lblPloicy(0)"
      Tab(0)Ctrl(4)Index=   0
      Tab(0)Ctrl(4)Name=   "lblPloicy"
      TabCaption1     =   "00540061006200200031"
      TabEnabled1     =   -1  'True
      TabColor1       =   8438015
      Tab(1)Controls  =   13
      Tab(1)Ctrl(1)ID =   "chkUse(1)"
      Tab(1)Ctrl(1)Index=   1
      Tab(1)Ctrl(1)Name=   "chkUse"
      Tab(1)Ctrl(2)ID =   "txtMSEncryptionType"
      Tab(1)Ctrl(2)Name=   "txtMSEncryptionType"
      Tab(1)Ctrl(3)ID =   "cmdBrowseCert(0)"
      Tab(1)Ctrl(3)Index=   0
      Tab(1)Ctrl(3)Name=   "cmdBrowseCert"
      Tab(1)Ctrl(4)ID =   "XPFrame302"
      Tab(1)Ctrl(4)Name=   "XPFrame302"
      Tab(1)Ctrl(5)ID =   "cboPolicy(1)"
      Tab(1)Ctrl(5)Index=   1
      Tab(1)Ctrl(5)Name=   "cboPolicy"
      Tab(1)Ctrl(6)ID =   "cboAlgorithm"
      Tab(1)Ctrl(6)Name=   "cboAlgorithm"
      Tab(1)Ctrl(7)ID =   "cboHash"
      Tab(1)Ctrl(7)Name=   "cboHash"
      Tab(1)Ctrl(8)ID =   "chkPGP"
      Tab(1)Ctrl(8)Name=   "chkPGP"
      Tab(1)Ctrl(9)ID =   "txtPGP(0)"
      Tab(1)Ctrl(9)Index=   0
      Tab(1)Ctrl(9)Name=   "txtPGP"
      Tab(1)Ctrl(10)ID=   "lblPloicy(1)"
      Tab(1)Ctrl(10)Index=   1
      Tab(1)Ctrl(10)Name=   "lblPloicy"
      Tab(1)Ctrl(11)ID=   "lblMSEncryptionType(0)"
      Tab(1)Ctrl(11)Index=   0
      Tab(1)Ctrl(11)Name=   "lblMSEncryptionType"
      Tab(1)Ctrl(12)ID=   "lblAlgorithm"
      Tab(1)Ctrl(12)Name=   "lblAlgorithm"
      Tab(1)Ctrl(13)ID=   "lblHash"
      Tab(1)Ctrl(13)Name=   "lblHash"
      TabCaption2     =   "00540061006200200032"
      TabEnabled2     =   -1  'True
      TabColor2       =   12640511
      Tab(2)Controls  =   4
      Tab(2)Ctrl(1)ID =   "chkAttachCert"
      Tab(2)Ctrl(1)Name=   "chkAttachCert"
      Tab(2)Ctrl(2)ID =   "cmdBrowseCert(1)"
      Tab(2)Ctrl(2)Index=   1
      Tab(2)Ctrl(2)Name=   "cmdBrowseCert"
      Tab(2)Ctrl(3)ID =   "cboPolicy(2)"
      Tab(2)Ctrl(3)Index=   2
      Tab(2)Ctrl(3)Name=   "cboPolicy"
      Tab(2)Ctrl(4)ID =   "lblPloicy(2)"
      Tab(2)Ctrl(4)Index=   2
      Tab(2)Ctrl(4)Name=   "lblPloicy"
      TabCaption3     =   "00540061006200200033"
      TabEnabled3     =   -1  'True
      TabColor3       =   12640511
      Tab(3)Controls  =   8
      Tab(3)Ctrl(1)ID =   "cboEncoding(3)"
      Tab(3)Ctrl(1)Index=   3
      Tab(3)Ctrl(1)Name=   "cboEncoding"
      Tab(3)Ctrl(2)ID =   "cboEncoding(0)"
      Tab(3)Ctrl(2)Index=   0
      Tab(3)Ctrl(2)Name=   "cboEncoding"
      Tab(3)Ctrl(3)ID =   "cboEncoding(1)"
      Tab(3)Ctrl(3)Index=   1
      Tab(3)Ctrl(3)Name=   "cboEncoding"
      Tab(3)Ctrl(4)ID =   "cboEncoding(2)"
      Tab(3)Ctrl(4)Index=   2
      Tab(3)Ctrl(4)Name=   "cboEncoding"
      Tab(3)Ctrl(5)ID =   "lblEncoding(3)"
      Tab(3)Ctrl(5)Index=   3
      Tab(3)Ctrl(5)Name=   "lblEncoding"
      Tab(3)Ctrl(6)ID =   "lblEncoding(0)"
      Tab(3)Ctrl(6)Index=   0
      Tab(3)Ctrl(6)Name=   "lblEncoding"
      Tab(3)Ctrl(7)ID =   "lblEncoding(1)"
      Tab(3)Ctrl(7)Index=   1
      Tab(3)Ctrl(7)Name=   "lblEncoding"
      Tab(3)Ctrl(8)ID =   "lblEncoding(2)"
      Tab(3)Ctrl(8)Index=   2
      Tab(3)Ctrl(8)Name=   "lblEncoding"
      TabCaption4     =   "00540061006200200034"
      TabEnabled4     =   -1  'True
      TabColor4       =   12640511
      Tab(4)Controls  =   20
      Tab(4)Ctrl(1)ID =   "chkPrivateKey"
      Tab(4)Ctrl(1)Name=   "chkPrivateKey"
      Tab(4)Ctrl(2)ID =   "XPStaticLabel11"
      Tab(4)Ctrl(2)Name=   "XPStaticLabel11"
      Tab(4)Ctrl(3)ID =   "txtPassphraseCert"
      Tab(4)Ctrl(3)Name=   "txtPassphraseCert"
      Tab(4)Ctrl(4)ID =   "cmbCSP"
      Tab(4)Ctrl(4)Name=   "cmbCSP"
      Tab(4)Ctrl(5)ID =   "cmbCertStoreLocation"
      Tab(4)Ctrl(5)Name=   "cmbCertStoreLocation"
      Tab(4)Ctrl(6)ID =   "XPStaticLabel1"
      Tab(4)Ctrl(6)Name=   "XPStaticLabel1"
      Tab(4)Ctrl(7)ID =   "XPStaticLabel2"
      Tab(4)Ctrl(7)Name=   "XPStaticLabel2"
      Tab(4)Ctrl(8)ID =   "cmbCertStoreName"
      Tab(4)Ctrl(8)Name=   "cmbCertStoreName"
      Tab(4)Ctrl(9)ID =   "txtCertSubjName"
      Tab(4)Ctrl(9)Name=   "txtCertSubjName"
      Tab(4)Ctrl(10)ID=   "dpFrom"
      Tab(4)Ctrl(10)Name=   "dpFrom"
      Tab(4)Ctrl(11)ID=   "dpTo"
      Tab(4)Ctrl(11)Name=   "dpTo"
      Tab(4)Ctrl(12)ID=   "btnGenerate"
      Tab(4)Ctrl(12)Name=   "btnGenerate"
      Tab(4)Ctrl(13)ID=   "btnView"
      Tab(4)Ctrl(13)Name=   "btnView"
      Tab(4)Ctrl(14)ID=   "XPStaticLabel3"
      Tab(4)Ctrl(14)Name=   "XPStaticLabel3"
      Tab(4)Ctrl(15)ID=   "XPStaticLabel4"
      Tab(4)Ctrl(15)Name=   "XPStaticLabel4"
      Tab(4)Ctrl(16)ID=   "XPStaticLabel5"
      Tab(4)Ctrl(16)Name=   "XPStaticLabel5"
      Tab(4)Ctrl(17)ID=   "XPStaticLabel6"
      Tab(4)Ctrl(17)Name=   "XPStaticLabel6"
      Tab(4)Ctrl(18)ID=   "XPStaticLabel7"
      Tab(4)Ctrl(18)Name=   "XPStaticLabel7"
      Tab(4)Ctrl(19)ID=   "txtIssuerName"
      Tab(4)Ctrl(19)Name=   "txtIssuerName"
      Tab(4)Ctrl(20)ID=   "btnExport"
      Tab(4)Ctrl(20)Name=   "btnExport"
      TabCaption5     =   "00540061006200200035"
      TabEnabled5     =   -1  'True
      TabColor5       =   12640511
      Tab(5)Controls  =   9
      Tab(5)Ctrl(1)ID =   "XPStaticLabel12"
      Tab(5)Ctrl(1)Name=   "XPStaticLabel12"
      Tab(5)Ctrl(2)ID =   "txtPrivateKey"
      Tab(5)Ctrl(2)Name=   "txtPrivateKey"
      Tab(5)Ctrl(3)ID =   "cbKeyType"
      Tab(5)Ctrl(3)Name=   "cbKeyType"
      Tab(5)Ctrl(4)ID =   "XPStaticLabel10"
      Tab(5)Ctrl(4)Name=   "XPStaticLabel10"
      Tab(5)Ctrl(5)ID =   "cbBitCount"
      Tab(5)Ctrl(5)Name=   "cbBitCount"
      Tab(5)Ctrl(6)ID =   "XPStaticLabel9"
      Tab(5)Ctrl(6)Name=   "XPStaticLabel9"
      Tab(5)Ctrl(7)ID =   "btnGeneratePair"
      Tab(5)Ctrl(7)Name=   "btnGeneratePair"
      Tab(5)Ctrl(8)ID =   "txtPassphrase"
      Tab(5)Ctrl(8)Name=   "txtPassphrase"
      Tab(5)Ctrl(9)ID =   "lblPassphrase"
      Tab(5)Ctrl(9)Name=   "lblPassphrase"
      TabCaption6     =   "00540061006200200036"
      TabEnabled6     =   -1  'True
      TabColor6       =   12640511
      Tab(6)Controls  =   6
      Tab(6)Ctrl(1)ID =   "XPStaticLabel13"
      Tab(6)Ctrl(1)Name=   "XPStaticLabel13"
      Tab(6)Ctrl(2)ID =   "txtPassPhraseImport"
      Tab(6)Ctrl(2)Name=   "txtPassPhraseImport"
      Tab(6)Ctrl(3)ID =   "txtImportFilePath"
      Tab(6)Ctrl(3)Name=   "txtImportFilePath"
      Tab(6)Ctrl(4)ID =   "btnViewCertificate"
      Tab(6)Ctrl(4)Name=   "btnViewCertificate"
      Tab(6)Ctrl(5)ID =   "btnImport"
      Tab(6)Ctrl(5)Name=   "btnImport"
      Tab(6)Ctrl(6)ID =   "lblArchiveFolder"
      Tab(6)Ctrl(6)Name=   "lblArchiveFolder"
      TabCaption7     =   "00540061006200200037"
      TabEnabled7     =   -1  'True
      TabColor7       =   12640511
      Tab(7)Controls  =   1
      Tab(7)Ctrl(1)ID =   "eplPrivateKey"
      Tab(7)Ctrl(1)Name=   "eplPrivateKey"
      TabCaption8     =   "00540061006200200038"
      TabEnabled8     =   -1  'True
      TabColor8       =   12640511
      Tab(8)Controls  =   4
      Tab(8)Ctrl(1)ID =   "txtPublicKey"
      Tab(8)Ctrl(1)Name=   "txtPublicKey"
      Tab(8)Ctrl(2)ID =   "XPStaticLabel14"
      Tab(8)Ctrl(2)Name=   "XPStaticLabel14"
      Tab(8)Ctrl(3)ID =   "txtRSA"
      Tab(8)Ctrl(3)Name=   "txtRSA"
      Tab(8)Ctrl(4)ID =   "XPStaticLabel8"
      Tab(8)Ctrl(4)Name=   "XPStaticLabel8"
      ClientAreaBackColor=   14737632
      TabSizing       =   1
      TabsPerRow      =   9
      CurrentTab      =   6
      CaptionAlignment=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      MousePointer    =   0
      Begin ciaXPSelectionControls30.XPCheckBox30 chkPrivateKey 
         Height          =   210
         Left            =   79500
         TabIndex        =   81
         Top             =   2820
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   370
         Caption         =   "0049006E0063006C007500640065002000500072006900760061007400650020004B0065007900200069006E0020004500780070006F00720074"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
         LicValid        =   -1  'True
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtPublicKey 
         Height          =   765
         Left            =   75390
         TabIndex        =   80
         Top             =   2850
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         Locked          =   -1  'True
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":20970
         MultiLine       =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel14 
         Height          =   210
         Left            =   75390
         TabIndex        =   79
         Top             =   2640
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Saved public key file:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel13 
         Height          =   210
         Left            =   450
         TabIndex        =   78
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Private Key Passphrase"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtPassPhraseImport 
         Height          =   765
         Left            =   420
         TabIndex        =   77
         Top             =   2160
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1349
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":2098C
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         RequiredField   =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel12 
         Height          =   210
         Left            =   75210
         TabIndex        =   76
         Top             =   2310
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Saved private key file:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtPrivateKey 
         Height          =   765
         Left            =   75210
         TabIndex        =   75
         Top             =   2520
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         Locked          =   -1  'True
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":209A8
         MultiLine       =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtRSA 
         Height          =   1425
         Left            =   75360
         TabIndex        =   74
         Top             =   1140
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2514
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":209C4
         MultiLine       =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel8 
         Height          =   210
         Left            =   75360
         TabIndex        =   73
         Top             =   900
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   370
         Caption         =   "RSA Public Key for pasting into OpenSSH Server"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin eLifePListL.ePropertyList eplPrivateKey 
         Height          =   2865
         Left            =   75300
         TabIndex        =   72
         Top             =   840
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   5054
         DescriptionBackColor=   16311512
         DescriptionForeColor=   -2147483630
         DescriptionForeColor=   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DescriptionHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ListItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DefaultBorderColor=   0   'False
         Headers         =   0   'False
         HeaderNameCaption=   ""
         HeaderValueCaption=   ""
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel11 
         Height          =   210
         Left            =   79500
         TabIndex        =   71
         Top             =   1950
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Private Key Passphrase"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtPassphraseCert 
         Height          =   615
         Left            =   79500
         TabIndex        =   70
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":209E0
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         RequiredField   =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cbKeyType 
         Height          =   300
         Left            =   79890
         TabIndex        =   69
         Top             =   1980
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel10 
         Height          =   210
         Left            =   79890
         TabIndex        =   68
         Top             =   1740
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Key Type:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPComboBox30.XPComboBox30 cbBitCount 
         Height          =   300
         Left            =   79860
         TabIndex        =   67
         Top             =   1380
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel9 
         Height          =   210
         Left            =   79860
         TabIndex        =   66
         Top             =   1140
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Bit Count:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPButton30.XPButton30 btnGeneratePair 
         Height          =   450
         Left            =   79920
         TabIndex        =   65
         Top             =   2610
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   794
         AutoSelectTheme =   -1  'True
         Caption         =   "00470065006E006500720061007400650020004B0065007900200050006100690072"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtPassphrase 
         Height          =   765
         Left            =   75180
         TabIndex        =   64
         Top             =   1380
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         ButtonMaskColor =   0
         ButtonCaptionVAlignment=   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":209FC
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         RequiredField   =   -1  'True
         ScrollBars      =   2
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel lblPassphrase 
         Height          =   210
         Left            =   75210
         TabIndex        =   63
         Top             =   1140
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Private Key Passphrase"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPText30.XPText30 txtImportFilePath 
         Height          =   345
         Left            =   420
         TabIndex        =   61
         Top             =   1410
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   609
         ButtonCaption   =   ""
         ButtonPicture   =   "Security.frx":20A18
         ButtonVisible   =   -1  'True
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnViewCertificate 
         Height          =   375
         Left            =   5850
         TabIndex        =   60
         Top             =   2910
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0056006900650077002000430065007200740069006600690063006100740065"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnImport 
         Height          =   375
         Left            =   5850
         TabIndex        =   59
         Top             =   2280
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         AutoSelectTheme =   -1  'True
         Caption         =   "0049006D0070006F00720074"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cmbCSP 
         Height          =   300
         Left            =   75240
         TabIndex        =   58
         Top             =   1050
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cmbCertStoreLocation 
         Height          =   300
         Left            =   79500
         TabIndex        =   57
         Top             =   1020
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel1 
         Height          =   210
         Left            =   75270
         TabIndex        =   56
         Top             =   840
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   370
         Caption         =   "Cryptographic Service Provider:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel2 
         Height          =   210
         Left            =   79500
         TabIndex        =   55
         Top             =   810
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Certificate store location:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPComboBox30.XPComboBox30 cmbCertStoreName 
         Height          =   300
         Left            =   79500
         TabIndex        =   54
         Top             =   1590
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPText30.XPText30 txtCertSubjName 
         Height          =   315
         Left            =   75240
         TabIndex        =   53
         Top             =   1590
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RequiredField   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPDatePicker30.XPDatePicker30 dpFrom 
         Height          =   285
         Left            =   75240
         TabIndex        =   52
         Top             =   2910
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         VisualStyle     =   2
         RequiredField   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatString    =   "mm/dd/yyyy"
         MouseIcon       =   "Security.frx":20DB2
         BorderColor     =   12164479
         DefaultBorderColors=   0   'False
         CurrentDate     =   "19/05/2007"
         BeginProperty CalendarButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarDayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarMonthFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarComboSpinFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPDatePicker30.XPDatePicker30 dpTo 
         Height          =   285
         Left            =   77160
         TabIndex        =   51
         Top             =   2910
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         VisualStyle     =   2
         RequiredField   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatString    =   "mm/dd/yyyy"
         MouseIcon       =   "Security.frx":20DCE
         BorderColor     =   12164479
         DefaultBorderColors=   0   'False
         CurrentDate     =   "19/05/2007"
         BeginProperty CalendarButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarDayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarMonthFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalendarComboSpinFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnGenerate 
         Height          =   510
         Left            =   80400
         TabIndex        =   50
         Top             =   3180
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   900
         AutoSelectTheme =   -1  'True
         Caption         =   "00470065006E0065007200610074006500200054006500730074002000430065007200740069006600690063006100740065"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnView 
         Height          =   300
         Left            =   75270
         TabIndex        =   49
         Top             =   3330
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         AutoSelectTheme =   -1  'True
         Caption         =   "0056006900650077002000430065007200740069006600690063006100740065"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel3 
         Height          =   210
         Left            =   75240
         TabIndex        =   48
         Top             =   1380
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Test Certiifcate Subject Name: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel4 
         Height          =   210
         Left            =   79500
         TabIndex        =   47
         Top             =   1380
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   370
         Caption         =   "Certificate store name:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel5 
         Height          =   210
         Left            =   75240
         TabIndex        =   46
         Top             =   2700
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   370
         Caption         =   "Valid From:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel6 
         Height          =   210
         Left            =   77160
         TabIndex        =   45
         Top             =   2700
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   370
         Caption         =   "Valid To:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPSideBarMenu30.XPStaticLabel XPStaticLabel7 
         Height          =   210
         Left            =   75240
         TabIndex        =   44
         Top             =   1980
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         Caption         =   "Certiifcate Issuer Name: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin ciaXPText30.XPText30 txtIssuerName 
         Height          =   315
         Left            =   75240
         TabIndex        =   43
         Top             =   2190
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboEncoding 
         Height          =   300
         Index           =   3
         Left            =   77850
         TabIndex        =   41
         Top             =   2220
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 btnExport 
         Height          =   300
         Left            =   77700
         TabIndex        =   37
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         AutoSelectTheme =   -1  'True
         Caption         =   "004500780070006F0072007400200043006500720074006900660069006300610074006500200074006F002000460069006C0065"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPSelectionControls30.XPCheckBox30 chkAttachCert 
         Height          =   315
         Left            =   75540
         TabIndex        =   27
         Top             =   1740
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         Caption         =   $"Security.frx":20DEA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 cmdBrowseCert 
         Height          =   465
         Index           =   1
         Left            =   78960
         TabIndex        =   26
         Top             =   2070
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         AutoSelectTheme =   -1  'True
         Caption         =   "005300690067006E006500720073"
         Picture         =   "Security.frx":20E7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPSelectionControls30.XPCheckBox30 chkUse 
         Height          =   210
         Index           =   1
         Left            =   76950
         TabIndex        =   25
         Top             =   2250
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   370
         Caption         =   "0048006900640065"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
         LicValid        =   -1  'True
      End
      Begin ciaXPFrame30.XPFrame30 XPFrame301 
         Height          =   1875
         Left            =   79260
         Top             =   1440
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   3307
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Radius          =   20
         LicValid        =   -1  'True
         Begin ciaXPSelectionControls30.XPCheckBox30 chkEncryptSign 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            ApplyCandyEffects=   -1  'True
            Caption         =   "0045006E00630072007900700074"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            LicValid        =   -1  'True
         End
         Begin ciaXPSelectionControls30.XPCheckBox30 chkEncryptSign 
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            Caption         =   "005300690067006E"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            LicValid        =   -1  'True
         End
      End
      Begin ciaXPMultiLineText30.XPMultiLineText30 txtMSEncryptionType 
         Height          =   525
         Left            =   76950
         TabIndex        =   22
         Top             =   2490
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   926
         ButtonMaskColor =   0
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledForeColor=   -2147483631
         DisabledBackColor=   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockedForeColor =   -2147483631
         LockedBackColor =   -2147483633
         LockedBackColor =   -2147483633
         LockedForeColor =   -2147483631
         MousePointer    =   0
         MouseIcon       =   "Security.frx":21218
         MultiLine       =   -1  'True
         ShowMenu        =   ""
         LicValid        =   -1  'True
      End
      Begin ciaXPButton30.XPButton30 cmdBrowseCert 
         Height          =   525
         Index           =   0
         Left            =   81600
         TabIndex        =   21
         Top             =   2490
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   926
         AutoSelectTheme =   -1  'True
         Picture         =   "Security.frx":21234
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPFrame30.XPFrame30 XPFrame302 
         Height          =   765
         Left            =   75570
         Top             =   1350
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1349
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Radius          =   20
         LicValid        =   -1  'True
         Begin ciaXPSelectionControls30.XPOption30 optMSEncryptionType 
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   360
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            Caption         =   "005000750062006C006900630020004B00650079"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            GroupID         =   6
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
         Begin ciaXPSelectionControls30.XPOption30 optMSEncryptionType 
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1860
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "00530079006D006D0065007400720069006300200065006E006300720079007000740069006F006E"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            GroupID         =   6
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
      End
      Begin ciaXPComboBox30.XPComboBox30 cboPolicy 
         Height          =   300
         Index           =   0
         Left            =   76110
         TabIndex        =   18
         Top             =   1350
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboPolicy 
         Height          =   300
         Index           =   1
         Left            =   75600
         TabIndex        =   17
         Top             =   960
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboPolicy 
         Height          =   300
         Index           =   2
         Left            =   75510
         TabIndex        =   16
         Top             =   1080
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboAlgorithm 
         Height          =   300
         Left            =   80490
         TabIndex        =   15
         Top             =   915
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboHash 
         Height          =   300
         Left            =   80490
         TabIndex        =   14
         Top             =   1260
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPFrame30.XPFrame30 fraPGPOptions 
         Height          =   2355
         Left            =   79770
         Top             =   600
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LicValid        =   -1  'True
         Begin ciaXPSelectionControls30.XPOption30 optPGP 
            Height          =   435
            Index           =   3
            Left            =   210
            TabIndex        =   10
            Top             =   1755
            Width           =   2040
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   $"Security.frx":215CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GroupID         =   7
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
         Begin ciaXPSelectionControls30.XPOption30 optPGP 
            Height          =   255
            Index           =   2
            Left            =   210
            TabIndex        =   11
            Top             =   1290
            Width           =   1755
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "005300690067006E006100740075007200650020004F006E006C0079"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GroupID         =   7
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
         Begin ciaXPSelectionControls30.XPOption30 optPGP 
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   12
            Top             =   810
            Width           =   2070
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "005000750062006C006900630020004B0065007900200065006E006300720079007000740069006F006E"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GroupID         =   7
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
         Begin ciaXPSelectionControls30.XPOption30 optPGP 
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   13
            Top             =   300
            Width           =   2070
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "0043006F006E00760065006E00740069006F006E0061006C00200065006E006300720079007000740069006F006E"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GroupID         =   7
            Grouped         =   -1  'True
            LicValid        =   -1  'True
         End
      End
      Begin ciaXPSelectionControls30.XPCheckBox30 chkPGP 
         Height          =   285
         Left            =   78960
         TabIndex        =   9
         Top             =   2100
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "004800690064006500200054007900700069006E0067"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         LicValid        =   -1  'True
      End
      Begin VB.TextBox txtPGP 
         Height          =   345
         Index           =   0
         Left            =   80430
         TabIndex        =   8
         Top             =   2010
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ciaXPComboBox30.XPComboBox30 cboEncoding 
         Height          =   300
         Index           =   0
         Left            =   77850
         TabIndex        =   7
         Top             =   1080
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboEncoding 
         Height          =   300
         Index           =   1
         Left            =   77850
         TabIndex        =   6
         Top             =   1470
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin ciaXPComboBox30.XPComboBox30 cboEncoding 
         Height          =   300
         Index           =   2
         Left            =   77850
         TabIndex        =   5
         Top             =   1860
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   16777215
         FullRowSelect   =   -1  'True
         LicValid        =   -1  'True
      End
      Begin VB.Label lblArchiveFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate File to Import: "
         Height          =   225
         Left            =   420
         TabIndex        =   62
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Label lblEncoding 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signature Encoding"
         Height          =   195
         Index           =   3
         Left            =   76365
         TabIndex        =   42
         Top             =   2250
         Width           =   1395
      End
      Begin VB.Label lblPloicy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ApplyTo"
         Height          =   225
         Index           =   2
         Left            =   75510
         TabIndex        =   36
         Top             =   870
         Width           =   645
      End
      Begin VB.Label lblPloicy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ApplyTo"
         Height          =   225
         Index           =   1
         Left            =   75600
         TabIndex        =   35
         Top             =   750
         Width           =   645
      End
      Begin VB.Label lblMSEncryptionType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encryption Certificate"
         Height          =   195
         Index           =   0
         Left            =   75300
         TabIndex        =   34
         Top             =   2430
         Width           =   1500
      End
      Begin VB.Label lblAlgorithm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Algorithm"
         Height          =   195
         Left            =   79740
         TabIndex        =   33
         Top             =   900
         Width           =   645
      End
      Begin VB.Label lblHash 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hash"
         Height          =   195
         Left            =   80010
         TabIndex        =   32
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label lblPloicy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cryptographic Engine"
         Height          =   225
         Index           =   0
         Left            =   76110
         TabIndex        =   31
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lblEncoding 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Encoding"
         Height          =   195
         Index           =   0
         Left            =   76410
         TabIndex        =   30
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Label lblEncoding 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header Encoding"
         Height          =   195
         Index           =   1
         Left            =   76485
         TabIndex        =   29
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label lblEncoding 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payload Encoding"
         Height          =   195
         Index           =   2
         Left            =   76470
         TabIndex        =   28
         Top             =   1890
         Width           =   1290
      End
   End
   Begin MSComDlg.CommonDialog cdDataFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' form    : frmSecurity
' DateTime  : 16/09/2005
' Purpose   : security related dialog to setup the properties of encryption and signature
' frames sizes:  Height:2920 ; Width: 4560; Top:825; Left:120;
'---------------------------------------------------------------------------------------
Public OrchestratorXML As COrchestXML


Public gstrLockID As String
Public gbSigned As Boolean
Public gbEncrypted As Boolean


Private mcapiCert As CAPICOM.Certificate
Private mwodCert As WODCERTMNGLib.Certificate

Private mProjectXML As CProjectStore

Private mbInSetup As Boolean
Private mbCancel As Boolean

Private mediDocument As Fredi.ediDocument
Private mediSecurities As Fredi.ediSecurities

Private Const mstrBLUECOLOR As Long = &HFFCDAB
Private Const mstrENGINE As String = "Engine"
Private Const mstrENCRIPTION As String = "Encryption"
Private Const mstrSIGNATURE As String = "Signature"
Private Const mstrENCODING As String = "Encoding"
Private Const mstrGENERATECERTIFICATE As String = "Generate"
Private Const mstrIMPORTCERTIFICATE As String = "Import"
Private Const mstrLOADKEYPAIR As String = "Load"
Private Const mstrPRIVATEKEY As String = "Private"
Private Const mstrPUPLICKEY As String = "Public"

Friend Property Get Canceled() As Boolean
    Canceled = mbCancel
End Property

Private Sub btnCancel_Click()
    
    On Error Resume Next
    mbCancel = True
    Unload Me
End Sub

Private Sub btnExport_Click()
    
    On Error GoTo ProcessError
    
    HourGlass Me
    Set mwodCert = New WODCERTMNGLib.Certificate
    If (txtPassphraseCert.Text <> "") Then
        'save also private key using passphrase
'        Set mcapiCert = FindCertificate(txtCertSubjName.Text)
        If Not mcapiCert Is Nothing Then
            If mcapiCert.HasPrivateKey Then
                If mcapiCert.PrivateKey.IsExportable Then
                    mcapiCert.Save gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".pfx", txtPassphraseCert.Text, CAPICOM_CERTIFICATE_SAVE_AS_PFX, CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN
                    DisplayStatus "Certificate (including private key) is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".pfx"
                    Sleep 3000
                    mwodCert.LoadKey gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".pfx", txtPassphraseCert.Text
                    txtRSA.Text = mwodCert.PublicKeyOpenSSH
                    txtPublicKey.Text = gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt"
                    PutText2File mwodCert.PublicKeyOpenSSH, gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt", True
                    DisplayStatus "Public key is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt"
                Else
                End If
            End If
        End If
    Else
        If Not mcapiCert Is Nothing Then
            If mcapiCert.HasPrivateKey Then
                If mcapiCert.PrivateKey.IsExportable Then
                    If (chkPrivateKey.Value = Checked) Then
                        Err.Raise 1, "Export Certificate", "You must specify a passphrase to export private key to file."
                    Else
                        mcapiCert.Save gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer", , CAPICOM_CERTIFICATE_SAVE_AS_CER, CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN
                        DisplayStatus "Certificate is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer"
                    End If
                Else
                    mcapiCert.Save gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer", , CAPICOM_CERTIFICATE_SAVE_AS_CER, CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN
                    DisplayStatus "Certificate is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer"
                End If
'                Sleep 3000
'                mwodCert.LoadKey gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer", txtPassphraseCert.Text
'                txtRSA.Text = mwodCert.PublicKeyOpenSSH
'                PutText2File mwodCert.PublicKeyOpenSSH, gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt", True
            Else
                mcapiCert.Save gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer", , CAPICOM_CERTIFICATE_SAVE_AS_CER, CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN
                DisplayStatus "Certificate is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".cer"
            End If
        Else
            Err.Raise 1, "Export Certificate", "You must specify load a certificate to export."
        End If
    End If
    HourGlass Me
    Exit Sub
ProcessError:
    MsgBox Err.Description, vbCritical, Err.Source
    HideStatus
    HourGlass Me
End Sub

Private Sub btnGeneratePair_Click()
    
    On Error GoTo ProcessError
    
    HourGlass Me
    Set mwodCert = New WODCERTMNGLib.Certificate
    If LenB(txtPassphrase.Text) = 0 Then
        Err.Raise 1, "Generate Key Pair", "You must have a passphrase to create key pair.", _
            vbCritical
    End If
    
    
    DisplayStatus "Move the mouse or press keyboard keys during generation for better results."
    Sleep 2000
    With mwodCert
        .BitCount = Val(cbBitCount.list(cbBitCount.ListIndex))
        .GenerateKey cbKeyType.ListIndex
        .SaveKey gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".pem", txtPassphrase.Text
        OrchestratorXML.PrivateKeyFile(gstrLockID) = gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".pem"
        eplPrivateKey.Properties.Clear
        DisplayStatus "Private key is exported to: " & vbCrLf & gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".pem"
        Sleep 3000
        txtRSA.Text = mwodCert.PublicKeyOpenSSH
        txtPublicKey.Text = gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".txt"
        PutText2File mwodCert.PublicKeyOpenSSH, gstrCodePath & "cert\" & Replace(gstrLockID, " ", "") & ".txt", True
        DisplayStatus "Public key is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(gstrLockID, " ", "") & ".txt"
    End With
    HourGlass Me
    Exit Sub
ProcessError:
    MsgBox Err.Description, vbCritical, Err.Source
    HourGlass Me
End Sub

Private Sub btnImport_Click()
    Dim Store As New CAPICOM.Store
    Dim tencryptEnvelop As TEncryptionEnvelope
    Dim tSign() As TCertificate

    ' Import test certificate.
    Dim oCertificate As Fredi.ediSecurityCertificate
    Dim fso As New Scripting.FileSystemObject
    
    
    On Error GoTo ProcessError
    
    If Not fso.FileExists(txtImportFilePath.Text) Then
        Err.Raise 1, , "Certificate or key pair file does not exists."
    End If
    HourGlass Me
'    Set oCertificate = mediSecurities.ImportCertificate(txtImportFilePath.Text)
    If OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
        Set mwodCert = New WODCERTMNGLib.Certificate
        mwodCert.LoadKey txtImportFilePath.Text, txtPassPhraseImport.Text
        DoEvents
        If mwodCert.HasPrivateKey Then
            txtRSA.Text = mwodCert.PublicKeyOpenSSH
            txtPublicKey.Text = gstrCodePath & "cert\" & Replace(gstrLockID, " ", "") & ".txt"
            PutText2File mwodCert.PublicKeyOpenSSH, gstrCodePath & "cert\" & Replace(gstrLockID, " ", "") & ".txt", True
            DisplayStatus "Public key is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(gstrLockID, " ", "") & ".txt"
            OrchestratorXML.PrivateKeyFile(gstrLockID) = txtImportFilePath.Text
            txtPrivateKey.Text = txtImportFilePath.Text
            OrchestratorXML.PassPhraseSSH(gstrLockID) = txtPassPhraseImport.Text
            txtPassphrase.Text = txtPassPhraseImport.Text
            eplPrivateKey.Properties.Clear
            tbsOptions.TabVisible(7) = True
            tbsOptions.TabVisible(8) = True
        Else
            tbsOptions.TabVisible(7) = False
            tbsOptions.TabVisible(8) = False
        End If
    Else
        Set mcapiCert = New CAPICOM.Certificate
        mcapiCert.Load txtImportFilePath.Text, txtPassPhraseImport.Text
        Store.Open CAPICOM_CURRENT_USER_STORE, , CAPICOM_STORE_OPEN_READ_WRITE
        Store.Load txtImportFilePath.Text, txtPassPhraseImport.Text
        If gbEncrypted Then
            tencryptEnvelop = OrchestratorXML.EnvEncryption(gstrLockID)
            tencryptEnvelop.tCer.strPassPhrase = txtPassPhraseImport.Text
            tencryptEnvelop.tCer.strSubject = Mid(mcapiCert.SubjectName, InStr(1, mcapiCert.SubjectName, "=") + 1)
            tencryptEnvelop.tCer.strIssuer = Mid(mcapiCert.IssuerName, InStr(1, mcapiCert.IssuerName, "=") + 1)
            OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
            txtCertSubjName.Text = tencryptEnvelop.tCer.strSubject
            txtIssuerName.Text = tencryptEnvelop.tCer.strIssuer
            txtMSEncryptionType.Text = tencryptEnvelop.tCer.strIssuer
        ElseIf gbSigned Then
            tSign = OrchestratorXML.Signature(gstrLockID)
            tSign(0).strPassPhrase = txtPassPhraseImport.Text
            tSign(0).strIssuer = Mid(mcapiCert.IssuerName, InStr(1, mcapiCert.IssuerName, "=") + 1)
            tSign(0).strSubject = Mid(mcapiCert.SubjectName, InStr(1, mcapiCert.SubjectName, "=") + 1)
            OrchestratorXML.Signature(gstrLockID) = tSign
        Else
            txtCertSubjName.Text = Mid(mcapiCert.SubjectName, InStr(1, mcapiCert.SubjectName, "=") + 1)
            txtIssuerName.Text = Mid(mcapiCert.IssuerName, InStr(1, mcapiCert.IssuerName, "=") + 1)
        End If
        OrchestratorXML.PrivateKeyFile(gstrLockID) = txtImportFilePath.Text
        OrchestratorXML.PassPhraseSSH(gstrLockID) = txtPassPhraseImport.Text
        txtPassphraseCert.Text = txtPassPhraseImport.Text
        btnViewCertificate.Enabled = True
        DoEvents
        If mcapiCert.HasPrivateKey Then
            Set mwodCert = New WODCERTMNGLib.Certificate
            mwodCert.LoadKey gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".pfx", txtPassPhraseImport.Text
            txtRSA.Text = mwodCert.PublicKeyOpenSSH
            txtPublicKey.Text = gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt"
            PutText2File mwodCert.PublicKeyOpenSSH, gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt", True
            DisplayStatus "Public key is exported to: " & vbCrLf & gstrCodePath & "cert\" & Replace(txtCertSubjName.Text, " ", "") & ".txt"
            eplPrivateKey.Properties.Clear
            tbsOptions.TabVisible(7) = True
            tbsOptions.TabVisible(8) = True
        Else
            tbsOptions.TabVisible(7) = False
            tbsOptions.TabVisible(8) = False
        End If
    End If
    eplPrivateKey.Properties.Clear
    HourGlass Me
    Exit Sub
ProcessError:
    MsgBox Err.Description, vbCritical, "Importing Certificate Failed"
    HourGlass Me, True
    
    

End Sub

Private Sub btnImport2Store_Click()
    Dim oCertLocs As Fredi.ediSecurityCertStoreLocations
    Dim oCertStores As Fredi.ediSecurityCertificateStores
    Dim oCertStore As Fredi.ediSecurityCertificateStore
    On Error Resume Next
    
    Set oCertLocs = mediSecurities.GetCertificateStoreLocations
    
    Set oCertStores = oCertLocs.GetCertificateStores(cmbCertStoreLocation.Text)
    Set oCertStore = oCertStores.GetCertificateStore(cmbCertStoreName.Text)
'    oCertStore.ImportCertificate btnImport2Store.Tag
    
End Sub

Private Sub btnInfo_Click()
    frmInfoText.lblInfoText.Caption = _
        "The certificate has been saved to the specified location. You can use this certificate to sign messages outbound to your partners or send the certificate to your partners for them to use it to encrypt messages they send to you."
    frmInfoText.Show vbModal
End Sub

Private Sub btnView_Click()
    
    On Error Resume Next

    
    If Not mcapiCert Is Nothing Then
        mcapiCert.display
    End If
End Sub

Private Sub btnViewCertificate_Click()
    
    On Error Resume Next

    If Not mcapiCert Is Nothing Then
        mcapiCert.display
    End If

End Sub

Private Sub cboEncoding_Click(Index As Integer)
    On Error Resume Next
    Dim eApply As EApplyTo
    
    If mbInSetup Then Exit Sub
    Select Case Index
        Case 0
            eApply = atCOMPLETE
        Case 1
            eApply = atHEADER
        Case 2
            eApply = atPAYLOAD
        Case 3
            eApply = atSIGNATURE
    End Select
    
    OrchestratorXML.SecurityEncoding(gstrLockID, eApply) = _
        OrchestratorXML.Text2Enum_EMessageEncoding(cboEncoding(Index).list(cboEncoding(Index).ListIndex))

End Sub

Private Sub cboPolicy_Change(Index As Integer)
    Dim bInSetup As Boolean
    On Error GoTo ProcessError

    If mbInSetup Then Exit Sub
    bInSetup = mbInSetup
    mbInSetup = True
    If Index = 0 Then
        OrchestratorXML.SecurityType(gstrLockID) = OrchestratorXML.Text2Enum_ESecurity(cboPolicy(0).list(cboPolicy(0).ListIndex))
        Select Case cboPolicy(0).list(cboPolicy(0).ListIndex)
            Case OrchestratorXML.Enum2Text_ESecurity(esPGP)
                LoadPGPInfo
            Case OrchestratorXML.Enum2Text_ESecurity(esMicrosoft)
                SecurityControls True
                SetupMSEncControls
                SetupMSSigControls
                OrchestratorXML.EncryptionApplyTo(gstrLockID) = atCOMPLETE
                OrchestratorXML.SignatureApplyTo(gstrLockID) = atCOMPLETE
                OrchestratorXML.RemoveEncryption gstrLockID
                tbsOptions.TabVisible(1) = True
                tbsOptions.TabVisible(2) = True
                tbsOptions.TabVisible(3) = True
                tbsOptions.TabVisible(4) = True
                tbsOptions.TabVisible(5) = False
                tbsOptions.TabVisible(7) = False
                tbsOptions.TabVisible(8) = False
                chkEncryptSign(0).Visible = True
                chkEncryptSign(1).Visible = True
            Case OrchestratorXML.Enum2Text_ESecurity(esSSH2)
                tbsOptions.TabVisible(1) = False
                tbsOptions.TabVisible(2) = False
                tbsOptions.TabVisible(3) = False
                tbsOptions.TabVisible(4) = False
                tbsOptions.TabVisible(5) = True
                tbsOptions.TabVisible(7) = False
                tbsOptions.TabVisible(8) = False
                chkEncryptSign(0).Visible = False
                chkEncryptSign(1).Visible = False
            Case Else
                SecurityControls False
        End Select
    ElseIf Index = 1 Then 'encryption apply to
        If gbEncrypted Then
            OrchestratorXML.EncryptionApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(1).list(cboPolicy(1).ListIndex))
            'if both encry and sign are checked then make both Applyto same
            If gbEncrypted And gbSigned Then
                cboPolicy(2).ListIndex = cboPolicy(1).ListIndex
                OrchestratorXML.SignatureApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(1).list(cboPolicy(1).ListIndex))
            End If
        End If
    ElseIf Index = 2 Then 'signature apply to
        If gbSigned Then
            OrchestratorXML.SignatureApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(cboPolicy(2).ListIndex))
            'if both encry and sign are checked then make both Applyto same
            If gbEncrypted And gbSigned Then
                cboPolicy(1).ListIndex = cboPolicy(2).ListIndex
                OrchestratorXML.EncryptionApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(cboPolicy(2).ListIndex))
            End If
        End If
    End If
    mbInSetup = bInSetup
    Exit Sub
ProcessError:


End Sub

Private Sub chkEncryptSign_Click(Index As Integer)
    Dim tCert As TCertificate
    Dim tSign() As TCertificate
    Dim emType As EMessagingType
    Dim tencryptEnvelop As TEncryptionEnvelope
    Dim est As ESecurityEngine
    
    On Error Resume Next

    If mbInSetup Then Exit Sub
    est = OrchestratorXML.SecurityType(gstrLockID)
    If Index = 0 Then
        If chkEncryptSign(0).Value = Checked Then
            gbEncrypted = True
            'add default certificate based security if exists
            tCert = mProjectXML.DefaultCertificate
            If est = esNONE Then
                OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft
            End If
            tencryptEnvelop.tCer.strIssuer = tCert.strIssuer
            tencryptEnvelop.tCer.strValidity = tCert.strValidity
            tencryptEnvelop.tCer.strID = tCert.strID
            OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
        Else
            gbEncrypted = False
            If OrchestratorXML.IsEncrypted(gstrLockID) Then
                OrchestratorXML.RemoveEncryption (gstrLockID)
            End If
        End If
    Else
        If chkEncryptSign(1).Value = Checked Then
            gbSigned = True
            'add default certificate based security if exists
            ReDim tSign(0)
            tSign(0) = mProjectXML.DefaultCertificate
            'check for messaging type
            If est = esNONE Then
                OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft
            End If
            OrchestratorXML.Signature(gstrLockID) = tSign
            cmdBrowseCert(1).Enabled = True
        Else
            gbSigned = False
            If OrchestratorXML.IsSigned(gstrLockID) Then
                OrchestratorXML.RemoveSignature gstrLockID
            End If
            cmdBrowseCert(1).Enabled = False
        End If
    End If
    LoadMSinfo
End Sub

Private Sub cmbCertStoreLocation_LostFocus()
    
    On Error Resume Next
    cmbCertStoreLocation_DropDown
    If cmbCertStoreLocation.Text <> cmbCertStoreLocation.list(cmbCertStoreLocation.FindItemIndex(cmbCertStoreLocation.Text, True)) Then
        cmbCertStoreLocation.ListIndex = 0
        cmbCertStoreLocation.Text = cmbCertStoreLocation.list(0)
    End If
End Sub

Private Sub cmbCertStoreName_LostFocus()
    On Error Resume Next
    cmbCertStoreName_DropDown
    If cmbCertStoreName.Text <> cmbCertStoreName.list(cmbCertStoreName.FindItemIndex(cmbCertStoreName.Text, True)) Then
        cmbCertStoreName.ListIndex = 0
        cmbCertStoreName.Text = cmbCertStoreName.list(0)
    End If

End Sub

Private Sub cmdOK_Click()
    
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%
    Dim astr() As String
    Dim enc As EEncoding
    Dim ato As EApplyTo
    Dim tee As TEncryptionEnvelope
    
    On Error Resume Next
    
    HourGlass Me
    
    mbCancel = False
    
    Me.Caption = "Security Settings - " & gstrLockID
    mbInSetup = True
    
    'load project.xml
    Set mProjectXML = New CProjectStore
    mProjectXML.LoadXML
    
    tbsOptions.TabCaption(0) = mstrENGINE
    tbsOptions.TabCaption(1) = mstrENCRIPTION
    tbsOptions.TabCaption(2) = mstrSIGNATURE
    tbsOptions.TabCaption(3) = mstrENCODING
    tbsOptions.TabCaption(4) = mstrGENERATECERTIFICATE
    tbsOptions.TabCaption(5) = mstrGENERATECERTIFICATE
    tbsOptions.TabCaption(6) = mstrIMPORTCERTIFICATE
    tbsOptions.TabCaption(7) = mstrPRIVATEKEY
    tbsOptions.TabCaption(8) = mstrPUPLICKEY
    
    Set tbsOptions.TabIcon(0) = FGlobalGraphics.img16.ItemPicture(riENGINE)
    Set tbsOptions.TabIcon(1) = FGlobalGraphics.img16.ItemPicture(riENCRYPT)
    Set tbsOptions.TabIcon(2) = FGlobalGraphics.img16.ItemPicture(riSIGN)
    Set tbsOptions.TabIcon(3) = FGlobalGraphics.img16.ItemPicture(riIMPORT)
    Set tbsOptions.TabIcon(4) = FGlobalGraphics.img16.ItemPicture(riCERTIFICATE)
    Set tbsOptions.TabIcon(5) = FGlobalGraphics.img16.ItemPicture(riKEY)
    Set tbsOptions.TabIcon(6) = FGlobalGraphics.img16.ItemPicture(riIMPORT)
    Set tbsOptions.TabIcon(7) = FGlobalGraphics.img16.ItemPicture(riKEY)
    Set tbsOptions.TabIcon(8) = FGlobalGraphics.img16.ItemPicture(riKEY)
    
    astr = OrchestratorXML.Options_ApplyTo
    
    
    'fill tab 0
    tbsOptions.CurrentTab = 0   'mstrENGINE
    If Not gbMSDrew Then CreateMS
    cboPolicy(0).Clear
    cboPolicy(0).AddItem OrchestratorXML.Enum2Text_ESecurity(esMicrosoft)
    cboPolicy(0).AddItem OrchestratorXML.Enum2Text_ESecurity(esSSH2)
    cboPolicy(0).AddItem OrchestratorXML.Enum2Text_ESecurity(esNONE)
    If OrchestratorXML.ExistsSecurity(gstrLockID) Then
        If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
            cboPolicy(0).ListIndex = 0 'microsoft
            If gbEncrypted Then
                chkEncryptSign(0).Value = Checked
            Else
                chkEncryptSign(0).Value = Unchecked
            End If
            If gbSigned Then
                chkEncryptSign(1).Value = Checked
            Else
                chkEncryptSign(1).Value = Unchecked
            End If
            chkEncryptSign(0).ValueSetClick = True
            tbsOptions.TabVisible(5) = False
        ElseIf OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
            tbsOptions.TabVisible(1) = False
            tbsOptions.TabVisible(2) = False
            tbsOptions.TabVisible(3) = False
            tbsOptions.TabVisible(4) = False
            cboPolicy(0).ListIndex = 1 'ssh2
            chkEncryptSign(0).Visible = False
            chkEncryptSign(1).Visible = False
        Else
            cboPolicy(0).ListIndex = 2 'none
        End If
    Else
        cboPolicy(0).ListIndex = 2 'none
    End If
    
    'fill tab 1
    'if the type was pgp then remove the signature and encryption
    chkUse(1).Value = vbChecked
    tbsOptions.CurrentTab = 1 'mstrENCRIPTION
    'hide all the children of this frame
    cboAlgorithm.Clear
    cboHash.Clear
    cboPolicy(1).Clear
    For i = 0 To UBound(astr)
        cboPolicy(1).AddItem astr(i)
    Next
    'apply to which part
    ato = OrchestratorXML.EncryptionApplyTo(gstrLockID)
    For i = 0 To cboPolicy(1).ListCount - 1
        If OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(1).list(i)) = ato Then
            cboPolicy(1).ListIndex = i
            Exit For
        End If
    Next i
    
    cboAlgorithm.Clear
    cboAlgorithm.AddItem gstrENCR_RC2
    cboAlgorithm.AddItem gstrENCR_RC4
    cboAlgorithm.AddItem gstrENCR_DES
    cboAlgorithm.AddItem gstrENCR_3DES

    cboHash.Clear
    cboHash.AddItem gstrHASH_SHA1
    cboHash.AddItem gstrHASH_MD2
    cboHash.AddItem gstrHASH_MD4
    cboHash.AddItem gstrHASH_MD5
    If gbEncrypted Then
        If OrchestratorXML.IsSymetric(gstrLockID) Then
            optMSEncryptionType(0).Value = True
            cboAlgorithm.Visible = True
            cboHash.Visible = True
            lblAlgorithm.Visible = True
            lblHash.Visible = True
            cboAlgorithm.Enabled = True
            cboHash.Enabled = True
            lblAlgorithm.Enabled = True
            lblHash.Enabled = True
            cmdBrowseCert(0).Enabled = False
            'for symmetric encryption, we can only allow base64,8bit and 7bit for encoding
            
            'rename password label and enlarge text box to allow for multiline input
            lblMSEncryptionType(0).Caption = "Password"
            txtMSEncryptionType.PasswordChar = "*"
        Else
            optMSEncryptionType(1).Value = True
            cboAlgorithm.Visible = False
            cboHash.Visible = False
            lblAlgorithm.Visible = False
            lblHash.Visible = False
            cmdBrowseCert(0).Enabled = True
            'when envelopped encryption, then we only allow base64 and binary
            'rename password label and enlarge text box to allow for multiline input
            lblMSEncryptionType(0).Caption = "Encryption Certificate"
            txtMSEncryptionType.Enabled = False
            txtMSEncryptionType.PasswordChar = ""
        End If
    End If
    If gbEncrypted Then 'if there is encryption
        If optMSEncryptionType(0).Value Then 'symetric
            Dim tsymencrypt As TEncryptionSymetric

            If OrchestratorXML.IsSymetric(gstrLockID) Then tsymencrypt = OrchestratorXML.SymmetricEncryption(gstrLockID)

            cboAlgorithm.Enabled = True
            cboHash.Enabled = True
            txtMSEncryptionType.Enabled = True
'
            txtMSEncryptionType.Text = tsymencrypt.strPassword
'
            For i = 0 To cboAlgorithm.ListCount - 1
                If OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboAlgorithm.list(i), esMicrosoft) = tsymencrypt.eAlgorithm Then
                    cboAlgorithm.ListIndex = i
                    Exit For
                End If
                cboAlgorithm.ListIndex = OrchestratorXML.Text2Enum_EEncryptionAlgorithm(gstrENCR_3DES, esMicrosoft)
            Next i
                        
            For i = 0 To cboHash.ListCount - 1
                If OrchestratorXML.Text2Enum_EHashAlgorithm(cboHash.list(i), esMicrosoft) = tsymencrypt.eHash Then
                    cboHash.ListIndex = i
                    Exit For
                End If
                cboHash.ListIndex = OrchestratorXML.Text2Enum_EHashAlgorithm(gstrHASH_SHA1, esMicrosoft)
            Next i
            
            
'
        ElseIf optMSEncryptionType(1).Value Then 'pk
            If Not OrchestratorXML.IsSymetric(gstrLockID) Then
                txtMSEncryptionType.Text = OrchestratorXML.EnvEncryption(gstrLockID).tCer.strIssuer
            Else
                txtMSEncryptionType.Text = ""
            End If
            txtMSEncryptionType.Enabled = False
        End If
        
    Else 'there is no encryption
        txtMSEncryptionType.Text = ""
        optMSEncryptionType(0).Enabled = False
        optMSEncryptionType(1).Enabled = False
    End If
    
    
    'fill tab 2
    tbsOptions.CurrentTab = 2       'mstrSIGNATURE
    cboPolicy(2).Clear
    For i = 0 To UBound(astr)
        cboPolicy(2).AddItem astr(i)
    Next
    'apply to which part
    ato = OrchestratorXML.SignatureApplyTo(gstrLockID)
    For i = 0 To cboPolicy(2).ListCount - 1
        If OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(i)) = ato Then
            cboPolicy(2).ListIndex = i
            Exit For
        End If
    Next i
    If gbSigned Then 'if there is signature
        Dim tSign() As TCertificate

        cmdBrowseCert(1).Enabled = True
        chkAttachCert.Value = IIf(OrchestratorXML.SignatureAttach(gstrLockID), 1, 0)
    Else    'there is no signature
        cmdBrowseCert(1).Enabled = False
    End If
    
    'fill tab 3
    tbsOptions.CurrentTab = 3   'mstrENCODING
    cboEncoding(0).Clear
    cboEncoding(0).AddItem gstr8BIT
    cboEncoding(0).AddItem gstr7BIT
    cboEncoding(0).AddItem gstrBASE64
    cboEncoding(0).AddItem gstrBINARY
'    cboEncoding(0).AddItem gstrBINARY
    enc = OrchestratorXML.SecurityEncoding(gstrLockID, atCOMPLETE)
    For i = 0 To cboEncoding(0).ListCount - 1
        If cboEncoding(0).list(i) = OrchestratorXML.Enum2Text_EMessageEncoding(enc) Then
            cboEncoding(0).ListIndex = i
            Exit For
        End If
    Next i
    
    cboEncoding(1).Clear
    cboEncoding(1).AddItem gstrBINARY
    cboEncoding(1).AddItem gstrBASE64
    cboEncoding(1).AddItem gstr8BIT
    cboEncoding(1).AddItem gstr7BIT
    enc = OrchestratorXML.SecurityEncoding(gstrLockID, atHEADER)
    For i = 0 To cboEncoding(1).ListCount - 1
        If cboEncoding(1).list(i) = OrchestratorXML.Enum2Text_EMessageEncoding(enc) Then
            cboEncoding(1).ListIndex = i
            Exit For
        End If
    Next i
    
    cboEncoding(2).Clear
    cboEncoding(2).AddItem gstrBINARY
    cboEncoding(2).AddItem gstrBASE64
    cboEncoding(2).AddItem gstr8BIT
    cboEncoding(2).AddItem gstr7BIT
    enc = OrchestratorXML.SecurityEncoding(gstrLockID, atPAYLOAD)
    For i = 0 To cboEncoding(2).ListCount - 1
        If cboEncoding(2).list(i) = OrchestratorXML.Enum2Text_EMessageEncoding(enc) Then
            cboEncoding(2).ListIndex = i
            Exit For
        End If
    Next i
    
    cboEncoding(3).Clear
    cboEncoding(3).AddItem gstrBINARY
    cboEncoding(3).AddItem gstrBASE64
    cboEncoding(3).AddItem gstrQP
    enc = OrchestratorXML.SecurityEncoding(gstrLockID, atSIGNATURE)
    For i = 0 To cboEncoding(3).ListCount - 1
        If cboEncoding(3).list(i) = OrchestratorXML.Enum2Text_EMessageEncoding(enc) Then
            cboEncoding(3).ListIndex = i
            Exit For
        End If
    Next i
    
    'fill tab 4
    tbsOptions.CurrentTab = 4   'mstrGENERATECERTIFICATE
    Set mediDocument = New Fredi.ediDocument
    Set mediSecurities = mediDocument.GetSecurities
    
    cmbCSP.Text = "Microsoft Strong Cryptographic Provider"
    cmbCertStoreLocation.Style = cboDropDownCombo
    cmbCertStoreLocation.Text = "CurrentUser"
    cmbCertStoreLocation.AutoComplete = True
    cmbCertStoreName.Text = "My"
    cmbCertStoreName.AutoComplete = True
    btnGenerate.Enabled = False
    btnView.Enabled = False
    btnImport.Enabled = False
    btnViewCertificate.Enabled = False
    btnExport.Enabled = False

    dpFrom.CurrentDate = Now()
    dpTo.CurrentDate = Now() + 365
    mediSecurities.DefaultProviderName = cmbCSP.Text
    mediSecurities.DefaultCertSystemStoreLocation = cmbCertStoreLocation.Text
    mediSecurities.DefaultCertSystemStoreName = cmbCertStoreName.Text
    If gbEncrypted Then
        tee = OrchestratorXML.EnvEncryption(gstrLockID)
        txtIssuerName.Text = tee.tCer.strIssuer
        txtCertSubjName.Text = tee.tCer.strSubject
        txtPassphraseCert.Text = tee.tCer.strPassPhrase
        If (txtCertSubjName.Text <> "") Then
            btnView.Enabled = True
            btnExport.Enabled = True
        End If
    ElseIf gbSigned Then 'if there is signature
        tSign = OrchestratorXML.Signature(gstrLockID)
        txtIssuerName.Text = tSign(0).strIssuer
        txtCertSubjName.Text = tSign(0).strSubject
        txtPassphraseCert.Text = tSign(0).strPassPhrase
        If (txtCertSubjName.Text <> "") Then
            btnView.Enabled = True
            btnExport.Enabled = True
        End If
    Else
        If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
            Set mcapiCert = New CAPICOM.Certificate
            mcapiCert.Load OrchestratorXML.PrivateKeyFile(gstrLockID), OrchestratorXML.PassPhraseSSH(gstrLockID)
            txtCertSubjName.Text = Mid(mcapiCert.SubjectName, InStr(1, mcapiCert.SubjectName, "=") + 1)
            txtIssuerName.Text = Mid(mcapiCert.IssuerName, InStr(1, mcapiCert.IssuerName, "=") + 1)
            txtPassphraseCert.Text = OrchestratorXML.PassPhraseSSH(gstrLockID)
        End If
    End If
    
    'fill tab 5
    'ssh2
    cbKeyType.AddItem "RSA Key"
    cbKeyType.AddItem "DSA Key"
    cbKeyType.ListIndex = 0
    cbBitCount.AddItem "768 bits"
    cbBitCount.AddItem "1024 bits"
    cbBitCount.AddItem "2048 bits"
    cbBitCount.AddItem "3072 bits"
    cbBitCount.ListIndex = 1
    txtPrivateKey.Text = OrchestratorXML.PrivateKeyFile(gstrLockID)
    txtPassphrase.Text = OrchestratorXML.PassPhraseSSH(gstrLockID)
    
    'fill tab 6
    tbsOptions.CurrentTab = 6   'mstrIMPORTCERTIFICATE
    
    'fill tab 7 private key
    Dim Priv As CAPICOM.PrivateKey
    Dim P As eLifePListL.Property20
    
'    eplPrivateKey.HoldDraw = True
    With eplPrivateKey
        .Headers = True
        .DescriptionVisible = True
        .AlternateRowColors = True
        .DividerPlacement = expLongestItem
        .ExpandingHeaderIcon = True
        .HeaderNameCaption = "Property Name"
        .HeaderValueCaption = "Value"
        .UseValueBackColors = True
        .BackColorEvenRows = RGB(247, 247, 247)
        .BackColorOddRows = RGB(234, 234, 234)
        .BackColorValue = RGB(248, 247, 239)
        .Properties.Clear
        .Style = Graphical
    End With
    
    If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
            Set mcapiCert = FindCertificate(txtCertSubjName.Text)
            If Not mcapiCert Is Nothing Then
                If mcapiCert.HasPrivateKey Then
                    Set Priv = mcapiCert.PrivateKey
                    With eplPrivateKey.Properties
                        .Clear
                        .Add "Accessible", PropertyTypeBoolean, Priv.IsAccessible
                        .Add "Exportable", PropertyTypeBoolean, Priv.IsExportable
                        .Add "Protected", PropertyTypeBoolean, Priv.IsProtected
    '                    .Add "Removable", PropertyTypeBoolean, Priv.IsRemovable
    '                    .Add "Hardware Device", PropertyTypeBoolean, Priv.IsHardwareDevice
    '                    .Add "Machine Keyset", PropertyTypeBoolean, Priv.IsMachineKeyset
                        Dim fso As New Scripting.FileSystemObject
                        Dim strFileName As String
                        strFileName = gstrCodePath & "cert\" & Trim(Replace(txtCertSubjName.Text, " ", "")) & ".pfx"
                        If Not fso.FileExists(strFileName) Then
                            .Add "Private Key File", PropertyTypeString, "*** Not created yet ***", , vbRed, , , "Click Generate tab, then click [Export Certificate to file] button"
                        Else
                            .Add "Private Key File", PropertyTypeString, strFileName
                        End If
                        .Add "Provider Name", PropertyTypeString, Priv.ProviderName
                        .Add "Container Name", PropertyTypeString, Priv.ContainerName
                        .Add "Unique Container Name", PropertyTypeString, Priv.UniqueContainerName
                    End With
                    'fill tab 4 generate tab
                    tbsOptions.CurrentTab = 4
                    If Priv.IsExportable Then
                        chkPrivateKey.Enabled = True
                        txtPassphraseCert.Enabled = True
                    Else
                        chkPrivateKey.Enabled = False
                        txtPassphraseCert.Enabled = False
                    End If
                    'fill tab 8 public key
                    tbsOptions.CurrentTab = 8
                    If fso.FileExists(strFileName) Then
                        Set mwodCert = New WODCERTMNGLib.Certificate
                        mwodCert.LoadKey strFileName, txtPassphraseCert.Text
                        txtRSA.Text = mwodCert.PublicKeyOpenSSH
                        strFileName = gstrCodePath & "cert\" & Trim(Replace(txtCertSubjName.Text, " ", "")) & ".txt"
                        If fso.FileExists(strFileName) Then
                            txtPublicKey.Text = gstrCodePath & "cert\" & Trim(Replace(txtCertSubjName.Text, " ", "")) & ".txt"
                        End If
                    End If
                Else
                    'fill tab 4 generate tab
                    tbsOptions.CurrentTab = 4
                    chkPrivateKey.Enabled = False
                    txtPassphraseCert.Enabled = False
                    tbsOptions.TabVisible(7) = False
                    tbsOptions.TabVisible(8) = False
                End If
            Else
                'fill tab 4 generate tab
                tbsOptions.CurrentTab = 4
                chkPrivateKey.Enabled = False
                txtPassphraseCert.Enabled = False
                tbsOptions.TabVisible(7) = False
                tbsOptions.TabVisible(8) = False
            End If
    ElseIf OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
        If (txtPrivateKey.Text <> "") Then
            Set mwodCert = New WODCERTMNGLib.Certificate
            mwodCert.LoadKey txtPrivateKey.Text, txtPassphrase.Text
            If Not mwodCert Is Nothing Then
                If mwodCert.HasPrivateKey Then
'                    Set Priv = mwodCert.PrivateKey
                    With eplPrivateKey.Properties
                        .Add "Accessible", PropertyTypeBoolean, "True"
                        .Add "Exportable", PropertyTypeBoolean, "True"
                        .Add "Protected", PropertyTypeBoolean, "True"
    '                    .Add "Removable", PropertyTypeBoolean, Priv.IsRemovable
    '                    .Add "Hardware Device", PropertyTypeBoolean, Priv.IsHardwareDevice
    '                    .Add "Machine Keyset", PropertyTypeBoolean, Priv.IsMachineKeyset
                        strFileName = txtPrivateKey.Text
                        If Not fso.FileExists(strFileName) Then
                            .Add "Private Key File", PropertyTypeString, "*** Not created yet ***", , vbRed, , , "Click Generate tab, then click [Export Certificate to file] button"
                        Else
                            .Add "Private Key File", PropertyTypeString, strFileName
                            Set mwodCert = New WODCERTMNGLib.Certificate
                            mwodCert.LoadKey txtPrivateKey.Text, txtPassphrase.Text
                            'fill tab 8 public key
                            tbsOptions.CurrentTab = 8
                            txtRSA.Text = mwodCert.PublicKeyOpenSSH
                            strFileName = gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".txt"
                            If fso.FileExists(strFileName) Then
                                txtPublicKey.Text = gstrCodePath & "cert\" & Trim(Replace(gstrLockID, " ", "")) & ".txt"
                            End If
                        End If
'                        .Add "Provider Name", PropertyTypeString, Priv.ProviderName
'                        .Add "Container Name", PropertyTypeString, Priv.ContainerName
'                        .Add "Unique Container Name", PropertyTypeString, Priv.UniqueContainerName
                    End With
                End If
            End If
        Else
            tbsOptions.TabVisible(7) = False
            tbsOptions.TabVisible(8) = False
        End If
    Else
        tbsOptions.TabVisible(7) = False
        tbsOptions.TabVisible(8) = False
    End If
    
    'fill tab 8 public key
'    If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
'        strFileName = gstrCodePath & "cert\" & Trim(Replace(txtCertSubjName.Text, " ", "")) & ".pfx"
'    ElseIf OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
'        If (txtPrivateKey.Text <> "") Then
'        End If
'    End If

    
    tbsOptions.CurrentTab = 0
    
    mbInSetup = False

    HourGlass Me


End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mediDocument.Close
    Set mediDocument = Nothing
    Set mediSecurities = Nothing

    Set OrchestratorXML = Nothing
    'unload objects that has used images
    Set tbsOptions.TabIcon(0) = Nothing
    Set tbsOptions.TabIcon(1) = Nothing
    Set tbsOptions.TabIcon(2) = Nothing
    Set tbsOptions.TabIcon(3) = Nothing
    Set tbsOptions.TabIcon(4) = Nothing
    Set tbsOptions.TabIcon(5) = Nothing
    Set tbsOptions.TabIcon(6) = Nothing
    Set tbsOptions.TabIcon(7) = Nothing
    Set tbsOptions.TabIcon(8) = Nothing
End Sub

'*********************************
' Returns the index of an item in a control
'*********************************
Private Function GetIndexOf(cb As Object, strText As String) As Integer
    Dim i As Integer
    
    For i = 0 To cb.ListCount - 1
        If cb.list(i) = strText Then
            GetIndexOf = i
            Exit Function
        End If
    Next i
    GetIndexOf = -1
End Function

Private Sub cboAlgorithm_Click()
    Dim tencryptsym As TEncryptionSymetric
            
    
    On Error Resume Next
    
    If mbInSetup Then Exit Sub
            
    'save in portxml
    tencryptsym.eAlgorithm = OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboAlgorithm.list(cboAlgorithm.ListIndex), esMicrosoft)
    tencryptsym.eHash = OrchestratorXML.Text2Enum_EHashAlgorithm(cboHash.list(cboHash.ListIndex), esMicrosoft)
    tencryptsym.strPassword = txtMSEncryptionType.Text
    OrchestratorXML.SymmetricEncryption(gstrLockID) = tencryptsym
    

End Sub

Private Sub cboHash_Click()
    Dim tencryptsym As TEncryptionSymetric
            
    
    On Error Resume Next
    
    If mbInSetup Then Exit Sub
    'save in portxml
    tencryptsym.eAlgorithm = OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboAlgorithm.list(cboAlgorithm.ListIndex), esMicrosoft)
    tencryptsym.eHash = OrchestratorXML.Text2Enum_EHashAlgorithm(cboHash.list(cboHash.ListIndex), esMicrosoft)
    tencryptsym.strPassword = txtMSEncryptionType.Text
    OrchestratorXML.SymmetricEncryption(gstrLockID) = tencryptsym
    

End Sub

Private Sub cboMSEncryptionType_click(Index As Integer)
End Sub

Private Sub cboPGP_click(Index As Integer)
    
    On Error Resume Next
    
End Sub
'
Private Sub cboPolicy_Click(Index As Integer)
    Dim bInSetup As Boolean
    On Error GoTo ProcessError

    If mbInSetup Then Exit Sub
    bInSetup = mbInSetup
    mbInSetup = True
    If Index = 0 Then
        OrchestratorXML.SecurityType(gstrLockID) = OrchestratorXML.Text2Enum_ESecurity(cboPolicy(0).list(cboPolicy(0).ListIndex))
        Select Case cboPolicy(0).Text
            Case OrchestratorXML.Enum2Text_ESecurity(esPGP)
                LoadPGPInfo
            Case OrchestratorXML.Enum2Text_ESecurity(esMicrosoft)
                If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then Exit Sub
                SecurityControls True
                SetupMSEncControls
                SetupMSSigControls
                OrchestratorXML.EncryptionApplyTo(gstrLockID) = atCOMPLETE
                OrchestratorXML.SignatureApplyTo(gstrLockID) = atCOMPLETE
                OrchestratorXML.RemoveEncryption gstrLockID
            Case OrchestratorXML.Enum2Text_ESecurity(esSSH2)
            Case Else
                SecurityControls False
        End Select
    ElseIf Index = 1 Then 'encryption apply to
        If gbEncrypted Then
            OrchestratorXML.EncryptionApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(1).list(cboPolicy(1).ListIndex))
            'if both encry and sign are checked then make both Applyto same
            If gbEncrypted And gbSigned Then
                cboPolicy(2).ListIndex = cboPolicy(1).ListIndex
                OrchestratorXML.SignatureApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(1).list(cboPolicy(1).ListIndex))
            End If
        End If
    ElseIf Index = 2 Then 'signature apply to
        If gbSigned Then
            OrchestratorXML.SignatureApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(cboPolicy(2).ListIndex))
            'if both encry and sign are checked then make both Applyto same
            If gbEncrypted And gbSigned Then
                cboPolicy(1).ListIndex = cboPolicy(2).ListIndex
                OrchestratorXML.EncryptionApplyTo(gstrLockID) = OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(cboPolicy(2).ListIndex))
            End If
        End If
    End If
    mbInSetup = bInSetup
    Exit Sub
ProcessError:

End Sub

Private Sub chkAttachCert_Click()
    
    On Error Resume Next
    
    If mbInSetup Then Exit Sub
    OrchestratorXML.SignatureAttach(gstrLockID) = IIf(chkAttachCert.Value = Checked, True, False)
'xpPopup_setup
'xpPopup_createMenus
'            xpPopup.PopupMenu "Signers", chkAttachCert.Left, chkAttachCert.Top
    
End Sub

Private Sub chkUse_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        Case 0 'encryption
            If gbEncrypted Then
                optMSEncryptionType(0).Enabled = True
                optMSEncryptionType(1).Enabled = True
                If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
                    SetupMSEncControls
                    OrchestratorXML.EncryptionApplyTo(gstrLockID) = atCOMPLETE
                ElseIf OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
                    LoadPGPInfo
                Else
                    Exit Sub
                End If
            Else 'we have to remove the values and disable all conrols
                OrchestratorXML.RemoveEncryption gstrLockID
                optMSEncryptionType(0).Enabled = False
                optMSEncryptionType(1).Enabled = False
                optMSEncryptionType(0).Value = False
                optMSEncryptionType(1).Value = False
                cboAlgorithm.Enabled = False
                cboHash.Enabled = False
                lblAlgorithm.Enabled = False
                lblHash.Enabled = False
                txtMSEncryptionType.Text = ""
            End If
        Case 1 'hide typing
            If chkUse(1).Value = vbChecked Then
                txtMSEncryptionType.PasswordChar = "*"
            Else
                txtMSEncryptionType.PasswordChar = ""
            End If
        Case 2 'signature
            If gbSigned Then
                Dim tCer() As TCertificate
                
                cmdBrowseCert(1).Enabled = True
                chkAttachCert.Value = Checked
                
                tCer(0) = mProjectXML.DefaultCertificate
                OrchestratorXML.Signature(gstrLockID) = tCer
                
                SetupMSSigControls
                OrchestratorXML.SignatureApplyTo(gstrLockID) = atCOMPLETE
            Else
                ' if no configuration was made, or it is empty
'                If txtMSEncryptionType(1).Text = "" Then Exit Sub
                OrchestratorXML.RemoveSignature gstrLockID
                cmdBrowseCert(1).Enabled = False
'                txtMSEncryptionType(1).Text = ""
            End If
        
    End Select
End Sub


Private Sub cmdBrowseCert_Click(Index As Integer)
    Dim atCertificate() As TCertificate
    
    On Error Resume Next
    
    Select Case Index
        Case 0 'encrypt
            Dim PickEncrypt As New fPickCert
            gbEncryption = True
            PickEncrypt.Left = Me.Left + Me.width \ 2 - PickEncrypt.width \ 2
            PickEncrypt.Top = Me.Top \ 2
            Me.WindowState = VBRUN.FormWindowStateConstants.vbNormal
        '            PickEncrypt.ZOrder 0
            PickEncrypt.Show vbModal
            atCertificate = PickEncrypt.Certificates
            txtMSEncryptionType.Text = atCertificate(0).strSubject & ": " & atCertificate(0).strIssuer ' GetCertificateInfo("encryption").strIssuer
            If txtMSEncryptionType.Text <> "" Then
                Dim tencryptEnvelop As TEncryptionEnvelope
                tencryptEnvelop.tCer.strIssuer = atCertificate(0).strIssuer
                tencryptEnvelop.tCer.strValidity = atCertificate(0).strValidity
                tencryptEnvelop.tCer.strID = atCertificate(0).strID
                tencryptEnvelop.tCer.strSubject = atCertificate(0).strSubject
                OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
                txtIssuerName.Text = atCertificate(0).strIssuer
                txtCertSubjName.Text = atCertificate(0).strSubject
                If (txtCertSubjName.Text <> "") Then
                    btnView.Enabled = True
                End If
                'renew variable
                Set mcapiCert = FindCertificate(txtCertSubjName.Text)
                eplPrivateKey.Properties.Clear
            End If
        Case 1 ' signature
            Dim PickSign As New fPickCert
            gbEncryption = False
            PickSign.Certificates = OrchestratorXML.Signature(gstrLockID)
            PickSign.Left = Me.Left + Me.width \ 2 - PickSign.width \ 2
            PickSign.Top = Me.Top \ 2
            Me.WindowState = VBRUN.FormWindowStateConstants.vbNormal
            PickSign.Show vbModal
            If PickSign.gbSave Then
                OrchestratorXML.Signature(gstrLockID) = PickSign.Certificates
                atCertificate = PickSign.Certificates
                txtIssuerName.Text = atCertificate(0).strIssuer
                txtCertSubjName.Text = atCertificate(0).strSubject
                If (txtCertSubjName.Text <> "") Then
                    btnView.Enabled = True
                End If
            End If
        Case 2 'assign keyrings
'            gstrKeyPortName = mrsTransports(xtfBOXNAME)
'            frmKeyrings.Show vbModal
'            If OrchestratorXML.getKeyRing(gstrLockID, False) <> "" And OrchestratorXML.getKeyRing(gstrLockID, True) <> "" Then
'                cmdBrowseCert(2).Caption = "Modify Keyrings"
''                lblPGP(5).Caption = "Assigned"
''                EnablePGPOptions True
'            Else
'                cmdBrowseCert(2).Caption = "Assign Keyrings"
''                lblPGP(5).Caption = "Not assigned"
''                EnablePGPOptions False
'            End If
        Case 3 'browse private key
'            fKeyList.Show vbModal
'            txtPGP(1).Text = gstrKeyId
        Case 4 'browse public key
'            fKeyList.Show vbModal
'            txtPGP(0).Text = gstrKeyId
    End Select
End Sub

Private Sub optMSEncryptionType_BeforeValueChanged(Index As Integer, state As Boolean, bCancel As Boolean)
'    Select Case Index
'        Case 0 'symmetric
'            'make signature available
'            'for symmetric encryption, we can only allow base64,8bit and 7bit for encoding
'            cboEncoding(0).Clear
'            cboEncoding(0).AddItem gstrBASE64
'            cboEncoding(0).AddItem gstr8BIT
'            cboEncoding(0).AddItem gstr7BIT
'            chkUse(1).Visible = True
'            chkUse(1).Enabled = True
'
'        Case 1 'public key
'            'if incoming case, no need to enter any
'            'certificate information.The coupler will search in the
'            'stores for a valid certificate to decrypt and validate
'            'the envelopped (public key encrypted) data.
'            chkUse(1).Visible = False
'            chkUse(1).Enabled = False
'            If Not mrsTransports(xtfINOUT) Then 'if incoming
''                fraMSEncryptionType(0).Visible = True
''                fraMSEncryptionType(1).Visible = False
'            Else
'                fraUse(1).Visible = True
''                fraMSEncryptionType(0).Visible = True
''                fraMSEncryptionType(1).Visible = False
'                'lets add a dummy value in port.xml file.
'                Dim tencryptEnvelop As TEncryptionEnvelope
'                tencryptEnvelop.tcer.strIssuer = "eLife Coupler"
'                tencryptEnvelop.tcer.strValidity = "01-01-1900"
'                tencryptEnvelop.tcer.strID = "Mr Dummy Fake"
'                OrchestratorXML.MS_AddPKEncryption tencryptEnvelop
'
'                mbSecurityChanged = True
'                EnableApplyCancel
'            End If
'            'hide signature
'            ' HK - 23-Jun-2004 10:54 v-4.0.1229
'            ' hussein kaddoura
'            '
'            'fraUse(1).Visible = False
'            'chkUse(2).Visible = False
'            'when envelopped encryption, then we only allow base64 and binary
'            cboEncoding(0).Clear
'            cboEncoding(0).AddItem gstrBASE64
'            cboEncoding(0).AddItem gstrBINARY
'
'
'    End Select

End Sub

Private Sub optMSEncryptionType_Click(Index As Integer)
    Dim tsymencrypt As TEncryptionSymetric
    Dim i%
    
    On Error Resume Next
    
    
    If mbInSetup Then Exit Sub
    Select Case Index
        Case 0 'symmetric
            cboAlgorithm.Visible = True
            cboHash.Visible = True
            lblAlgorithm.Visible = True
            lblHash.Visible = True
            chkUse(1).Visible = True
            
            'for symmetric encryption, we can only allow base64,8bit and 7bit for encoding
            cboEncoding(0).Clear
            cboEncoding(0).AddItem gstrBASE64
            cboEncoding(0).AddItem gstr8BIT
            cboEncoding(0).AddItem gstr7BIT
            cboEncoding(0).AddItem gstrQP
            cboEncoding(0).ListIndex = 0
            
            cboEncoding(1).Clear
            cboEncoding(1).AddItem gstrBASE64
            cboEncoding(1).AddItem gstr8BIT
            cboEncoding(1).AddItem gstr7BIT
            cboEncoding(1).AddItem gstrQP
            cboEncoding(1).ListIndex = 0
            
            cboEncoding(2).Clear
            cboEncoding(2).AddItem gstrBASE64
            cboEncoding(2).AddItem gstr8BIT
            cboEncoding(2).AddItem gstr7BIT
            cboEncoding(2).AddItem gstrQP
            cboEncoding(2).ListIndex = 0
            
            'rename password label and enlarge text box to allow for multiline input
            lblMSEncryptionType(0).Caption = "Password"
            txtMSEncryptionType.PasswordChar = "*"
            txtMSEncryptionType.Text = "eLife"
            txtMSEncryptionType.Enabled = True
            For i = 0 To cboAlgorithm.ListCount - 1
                If OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboAlgorithm.list(i), esMicrosoft) = eaCAPICOM_3DES Then
                    cboAlgorithm.ListIndex = i
                    Exit For
                End If
            Next i
                        
            For i = 0 To cboHash.ListCount - 1
                If OrchestratorXML.Text2Enum_EHashAlgorithm(cboHash.list(i), esMicrosoft) = haCAPICOM_SHA1 Then
                    cboHash.ListIndex = i
                    Exit For
                End If
            Next i
            
            'save in portxml
            tsymencrypt.eAlgorithm = eaCAPICOM_3DES
            tsymencrypt.eHash = haCAPICOM_SHA1
            tsymencrypt.strPassword = "eLife"
            OrchestratorXML.SymmetricEncryption(gstrLockID) = tsymencrypt
            
        Case 1 'public key
            'if incoming case, no need to enter any
            'certificate information.The coupler will search in the
            'stores for a valid certificate to decrypt and validate
            'the envelopped (public key encrypted) data.
            Dim tencryptEnvelop As TEncryptionEnvelope
            Dim tCer As TCertificate
            
            cboAlgorithm.Visible = False
            cboHash.Visible = False
            lblAlgorithm.Visible = False
            lblHash.Visible = False
            chkUse(1).Visible = False
            
            tCer = mProjectXML.DefaultCertificate
            tencryptEnvelop.tCer.strIssuer = tCer.strIssuer
            tencryptEnvelop.tCer.strValidity = tCer.strValidity
            tencryptEnvelop.tCer.strID = tCer.strID
            OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
            
            'rename password label and enlarge text box to allow for multiline input
            lblMSEncryptionType(0).Caption = "Encryption Certificate"
            txtMSEncryptionType.PasswordChar = ""
            txtMSEncryptionType.Enabled = False
            txtMSEncryptionType.Text = tCer.strIssuer

    End Select
End Sub
'
'Private Sub optPGP_Click(Index As Integer)
'    Dim tpgpsym As TEncryptionSymetric
'    Dim tpgpsign As TSignaturePK
'    Dim tpgpPK As TEncryptionPK
'
'    On Error GoTo ProcessError
'    'tpgpencrypt = getPGPEncryption
'    'tpgpsignature = getPGPSignature
'    txtPGP(0).Text = ""
'    txtPGP(1).Text = ""
'    txtPGP(2).Text = ""
'    cboPGP(0).ListIndex = -1
'    cboPGP(1).ListIndex = -1
'
'    Select Case Index
'        Case 0 'conventional
'            ShowPGPConventional
'            If OrchestratorXML.SecurityType = esPGP Then tpgpsym = OrchestratorXML.SymmetricEncryption
'            txtPGP(2).Text = tpgpsym.strPassword
'            If Not mrsTransports(xtfINOUT) Then
'                cboPGP(0).ListIndex = tpgpsym.eAlgorithm
'            End If
'        Case 1 ' public key
'            If OrchestratorXML.SecurityType = esPGP Then tpgpPK = OrchestratorXML.GetPKEncryption
'            If mrsTransports(xtfINOUT) Then 'incoming
'                ShowPGPInPublicKey
'                txtPGP(2).Text = tpgpPK.strPassPhrase
'            Else
'                ShowPGPOutPublicKey
'                txtPGP(0).Text = tpgpPK.strPublicKey
'            End If
'        Case 2 'signature
'            If OrchestratorXML.SecurityType = esPGP Then tpgpsign = OrchestratorXML.getPKSignature
'            If mrsTransports(xtfINOUT) Then 'incoming
'                ShowPGPInSignature
'            Else
'                ShowPGPOutSignature
'                txtPGP(1).Text = tpgpsign.strPrivateKey
'                txtPGP(2).Text = tpgpsign.strPassPhrase
'                cboPGP(1).ListIndex = tpgpsign.eHash
'            End If
'        Case 3 'both
'            If OrchestratorXML.SecurityType = esPGP Then
'                tpgpsign = OrchestratorXML.getPKSignature
'                tpgpPK = OrchestratorXML.GetPKEncryption
'            End If
'            If mrsTransports(xtfINOUT) Then 'incoming
'                ShowPGPInBoth
'                txtPGP(2).Text = tpgpPK.strPassPhrase
'            Else
'                ShowPGPOutBoth
'                txtPGP(1).Text = tpgpsign.strPrivateKey
'                txtPGP(2).Text = tpgpsign.strPassPhrase
'                cboPGP(1).ListIndex = tpgpsign.eHash
'                txtPGP(0).Text = tpgpPK.strPublicKey
'            End If
'    End Select
'    mbPGPEncryptionChanged = True
'    EnableApplyCancel
'    Exit Sub
'ProcessError:
'
'End Sub
'
'Private Sub optSecurityPolicy_Click(Index As Integer)
'
'
'    Select Case Index
'        Case 0 'none
''            fraSecurityType(0).Visible = False
''            fraSecurityType(1).Visible = False
'            mbEncryptionChanged = False
'            mbPGPEncryptionChanged = False
'            mbNoneActivated = True
'            EnableApplyCancel
'        Case 1 'microsoft
'            LoadMSinfo
''            fraSecurityType(0).Visible = True
''            fraSecurityType(1).Visible = False
'            mbPGPEncryptionChanged = False
'            mbNoneActivated = False
'            EnableApplyCancel
'        Case 2 'pgp
'            LoadPGPInfo
''            fraSecurityType(0).Visible = False
''            fraSecurityType(1).Visible = True
'            mbEncryptionChanged = False
'            mbNoneActivated = False
'            EnableApplyCancel
'    End Select
'End Sub

Private Sub tbsOptions_Click(ByVal CurrentTab As Integer, ByVal PreviousTab As Integer)
    Dim xElement As MSXML2.IXMLDOMNode
    Dim intMessageFormat As Integer
    Dim fra  As ciaXPFrame30.XPFrame30

    
    On Error Resume Next
    If mbInSetup = True Then Exit Sub
    HideStatus
    mbInSetup = True
    Select Case CurrentTab
        Case 0 'mstrENGINE
            lblTitle.Caption = "Set Cryptographic Engine"
            lblComment.Caption = "Select the cryptographic engine from the available list. You must insure that the selected engine is installed on your computer. In case of Microsoft engine, windows comes with that engine."
        Case 2 'mstrSIGNATURE
            lblTitle.Caption = "Set Signature Parameters"
            lblComment.Caption = "Signatures uses digital IDs or certificates to sign a message. You can co-sign a message by selecting multiple certificates."
        Case 1 'mstrENCRIPTION
            lblTitle.Caption = "Set Encryption Parameters"
            lblComment.Caption = "Encryption of a message can be done in two ways: using public key or using password. When using public key, the Coupler will use a certificate to encrypt the message. Password encryption is used when both party use common password."
        Case 3 'mstrENCODING
            lblTitle.Caption = "Set Message Encoding"
            lblComment.Caption = "Encrypted and/or signed Messages are normally transmitted in binary format, however, you can use base64 encoding."
        Case 4 'mstrGENERATECERTIFICATE
            lblTitle.Caption = "Generate Certificate"
            lblComment.Caption = "Creates your private certificates for in-house use or testing. Do not use these certificates with your trading partners."
            cmbCertStoreName_DropDown
            cmbCertStoreLocation_DropDown
        Case 5 'mstrGENERATECERTIFICATE
            lblTitle.Caption = "Generate Key Pair"
            lblComment.Caption = "Create SSH key pair."
            txtPrivateKey.Text = OrchestratorXML.PrivateKeyFile(gstrLockID)
        Case 6 'mstrIMPORTCERTIFICATE
            lblTitle.Caption = "Import Certificate"
            lblComment.Caption = "Import certifcates of your partners to verify their signatures or import your own certificates that you received from trusted authorities to sign or encrypt your documents."
        Case 7
            lblTitle.Caption = "Private Key Properties"
            lblComment.Caption = ""
        Case 8
            lblTitle.Caption = "Public Key Properties"
            lblComment.Caption = "Import certifcates of your partners to verify their signatures or import your own certificates that you received from trusted authorities to sign or encrypt your documents."
    End Select

    mbInSetup = False

End Sub

Private Sub tbsOptions_TabChanged(PreviousTab As Integer, CurrentTab As Integer)
    Dim Priv As CAPICOM.PrivateKey
    Dim P As eLifePListL.Property20

    On Error Resume Next
    
    If mbInSetup = True Then Exit Sub
    HideStatus
    Select Case CurrentTab
        Case 7 'private key
            If eplPrivateKey.Properties.Count <> 0 Then Exit Sub
            If OrchestratorXML.SecurityType(gstrLockID) = esMicrosoft Then
                    If Not mcapiCert Is Nothing Then
                        If mcapiCert.HasPrivateKey Then
                            Set Priv = mcapiCert.PrivateKey
                            With eplPrivateKey.Properties
                                .Clear
                                .Add "Accessible", PropertyTypeBoolean, Priv.IsAccessible
                                .Add "Exportable", PropertyTypeBoolean, Priv.IsExportable
                                .Add "Protected", PropertyTypeBoolean, Priv.IsProtected
            '                    .Add "Removable", PropertyTypeBoolean, Priv.IsRemovable
            '                    .Add "Hardware Device", PropertyTypeBoolean, Priv.IsHardwareDevice
            '                    .Add "Machine Keyset", PropertyTypeBoolean, Priv.IsMachineKeyset
                                Dim fso As New Scripting.FileSystemObject
                                Dim strFileName As String
                                strFileName = gstrCodePath & "cert\" & Trim(Replace(txtCertSubjName.Text, " ", "")) & ".pfx"
                                If Not fso.FileExists(strFileName) Then
                                    .Add "Private Key File", PropertyTypeString, "*** Not created yet ***", , vbRed, , , "Click Generate tab, then click [Export Certificate to file] button"
                                Else
                                    .Add "Private Key File", PropertyTypeString, strFileName
                                    If fso.FileExists(strFileName) Then
                                        Set mwodCert = New WODCERTMNGLib.Certificate
                                        mwodCert.LoadKey strFileName, txtPassphraseCert.Text
                                    End If
                                End If
                                .Add "Provider Name", PropertyTypeString, Priv.ProviderName
                                .Add "Container Name", PropertyTypeString, Priv.ContainerName
                                .Add "Unique Container Name", PropertyTypeString, Priv.UniqueContainerName
                            End With
                        End If
                    Else
                        tbsOptions.TabVisible(7) = False
                        tbsOptions.TabVisible(8) = False
                    End If
            ElseIf OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
                If (txtPrivateKey.Text <> "") Then
                    If Not mwodCert Is Nothing Then
                        If mwodCert.HasPrivateKey Then
        '                    Set Priv = mwodCert.PrivateKey
                            With eplPrivateKey.Properties
                                .Clear
                                .Add "Accessible", PropertyTypeBoolean, "True"
                                .Add "Exportable", PropertyTypeBoolean, "True"
                                .Add "Protected", PropertyTypeBoolean, "True"
            '                    .Add "Removable", PropertyTypeBoolean, Priv.IsRemovable
            '                    .Add "Hardware Device", PropertyTypeBoolean, Priv.IsHardwareDevice
            '                    .Add "Machine Keyset", PropertyTypeBoolean, Priv.IsMachineKeyset
                                strFileName = txtPrivateKey.Text
                                If Not fso.FileExists(strFileName) Then
                                    .Add "Private Key File", PropertyTypeString, "*** Not created yet ***", , vbRed, , , "Click Generate tab, then click [Export Certificate to file] button"
                                Else
                                    .Add "Private Key File", PropertyTypeString, strFileName
                                End If
        '                        .Add "Provider Name", PropertyTypeString, Priv.ProviderName
        '                        .Add "Container Name", PropertyTypeString, Priv.ContainerName
        '                        .Add "Unique Container Name", PropertyTypeString, Priv.UniqueContainerName
                            End With
                        End If
                    End If
                Else
                    tbsOptions.TabVisible(7) = False
                    tbsOptions.TabVisible(8) = False
                End If
            Else
                tbsOptions.TabVisible(7) = False
                tbsOptions.TabVisible(8) = False
            End If
        Case 8 'public key tab
            If Not mwodCert Is Nothing Then
                txtRSA.Text = mwodCert.PublicKeyOpenSSH
            Else
                txtRSA.Text = ""
            End If
        Case 4 'generate tab
            If Not mcapiCert Is Nothing Then
                If mcapiCert.HasPrivateKey Then
                    Set Priv = mcapiCert.PrivateKey
                    If Priv.IsExportable Then
                        chkPrivateKey.Enabled = True
                        txtPassphraseCert.Enabled = True
                    Else
                        chkPrivateKey.Enabled = False
                        txtPassphraseCert.Enabled = False
                    End If
                Else
                    chkPrivateKey.Enabled = False
                    txtPassphraseCert.Enabled = False
                End If
            End If
        Case Else
    End Select
End Sub

Private Sub txtImportFilePath_ButtonClick()
    Dim fso As New Scripting.FileSystemObject
    On Error Resume Next
    
    If fso.FileExists(txtImportFilePath.Text) Then
        cdDataFile.InitDir = fso.GetFile(txtImportFilePath.Text).ParentFolder
    ElseIf fso.FolderExists(txtImportFilePath.Text) Then
        cdDataFile.InitDir = txtImportFilePath.Text
    Else
        cdDataFile.InitDir = gstrCodePath & gstrPROJECTS & "\" & gstrProjectPath
    End If
    cdDataFile.CancelError = True
    If OrchestratorXML.SecurityType(gstrLockID) = esSSH2 Then
        cdDataFile.DialogTitle = "Select the Private Key File to import"
        cdDataFile.Filter = "All (*.ppk;*.pem;*.pfx)|*.ppk;*.pem;*.pfx|PuttyGen (*.ppk)|*.ppk|Privacy Enhanced Mail (*.pem)|*.pem|PKCS#12 format (*.pfx)|*.pfx"
    Else
        cdDataFile.DialogTitle = "Select the Certificate File to import"
        cdDataFile.Filter = "All (*.cer;*.p7b;*.pfx)|*.cer;*.p7b;*.pfx|PKCS#12 format (*.pfx)|*.pfx|DER or Base64 Encoded X.509 *.cer|*.cer|Crypto Message Syntax Std. PKCS#7 *.p7b|*.p7b"
    End If
    'we shouldn't allow multiple selection for now.
    cdDataFile.flags = cdlOFNExplorer Or cdlOFNHideReadOnly
    cdDataFile.ShowOpen
    If Err <> 0 Then
        Exit Sub
    End If
    txtImportFilePath.Text = cdDataFile.FileName
    If LenB(Trim(txtImportFilePath.Text)) <> 0 Then
        btnImport.Enabled = True
    End If
    Exit Sub

End Sub

Private Sub txtCertSubjName_Change()
    If mbInSetup Then Exit Sub
    If LenB(txtCertSubjName.Text) <> 0 Then
        btnGenerate.Enabled = True
        btnExport.Enabled = True
        chkPrivateKey.Enabled = True
        txtPassphraseCert.Enabled = True
    End If
'    txtIssuerName.Text = txtCertSubjName.Text
'    txtIssuerName.Locked = True
End Sub

Private Sub txtImportFilePath_Change()
    If LenB(Trim(txtImportFilePath.Text)) <> 0 Then
        btnImport.Enabled = True
    End If

End Sub

Private Sub txtMSEncryptionType_LostFocus()
    
    On Error Resume Next

    If optMSEncryptionType(0).Value Then
        'symmetric
        Dim tencryptsym As TEncryptionSymetric
        If txtMSEncryptionType.Text <> "" Then
            tencryptsym.strPassword = txtMSEncryptionType.Text 'mcCrypto.MSEncrypt(txtMSEncryptionType(1).Text, "", True, frezStreamEncryption)
        Else
            tencryptsym.strPassword = ""
        End If
        tencryptsym.eAlgorithm = OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboAlgorithm.list(cboAlgorithm.ListIndex), esMicrosoft)
        tencryptsym.eHash = OrchestratorXML.Text2Enum_EHashAlgorithm(cboHash.list(cboHash.ListIndex), esMicrosoft)
        OrchestratorXML.SymmetricEncryption(gstrLockID) = tencryptsym
    Else
        'public key cerificate changed is handled by certificate browse
    End If
        

End Sub
'*****************************************************
'this sub will create the microsoft security controls.
'******************************************************
Private Sub CreateMS()
        
        optMSEncryptionType(1).Visible = True
        'setting up the PK frame
        lblMSEncryptionType(0).Visible = True
        
        txtMSEncryptionType.Visible = True
        
        chkUse(1).Visible = True
        
        lblAlgorithm.Visible = True
        
        cboAlgorithm.Visible = True
        
        lblAlgorithm.Visible = True
        
        cboHash.Visible = True

        'setting up the signature frame
        gbMSDrew = True
End Sub
'
''*****************************************************
''this sub will create the PGP security controls.
''******************************************************
''
'Private Sub CreatePGP()
'    '1)The control limit on the form has been reached.that's why we have to create
'        'them on run-time
'    '2)Create the controls
'    '3)define the control containers
'    '4)position each control inside its container
'    '5)make visible
'    '6)set the first time creation flag to false so we only create them once.
'
'    '2)
'        'lblpgp(0) 'recipient KeyId
'        'txtPGP(0)      'recipient keyId field
'        'cbopgp(0)  'algorithm selection
''        Load fraSecurityType(1) ' pgp frame
'        Load lblPGP(1) 'private keyid
'        Load lblPGP(2) 'passphrase
'        Load lblPGP(3) 'hash
'        Load lblPGP(4) 'algorithm
'        Load lblPGP(5)          'assigned or not.
'        Load lblPGP(6) 'signature extension
'        Load txtPGP(1) 'private keyId field
'        Load txtPGP(2) 'passphrase field
'        Load txtPGP(3) 'signature extension
'        Load cboPGP(1) 'hashing selection
'        Load cmdBrowseCert(2) 'assign keyrings.
'        Load cmdBrowseCert(3) 'list the private keys
'        Load cmdBrowseCert(4) 'list the public keys
'    '3)
''        Set fraSecurityType(1).Container = fraTransactionSecurity
''        Set fraPGPOptions.Container = fraSecurityType(1)
''        Set cmdBrowseCert(2).Container = fraSecurityType(1)
''        Set cmdBrowseCert(3).Container = fraSecurityType(1)
''        Set cmdBrowseCert(4).Container = fraSecurityType(1)
'        'Set lblPGP(5).Container = fraSecurityType(1)
'
''        Dim i As Integer
''        For i = 0 To 6
''            Set lblPGP(i).Container = fraSecurityType(1)
''        Next i
''        For i = 0 To 3
''            Set txtPGP(i).Container = fraSecurityType(1)
''        Next i
''        Set cboPGP(0).Container = fraSecurityType(1)
''        Set cboPGP(1).Container = fraSecurityType(1)
''        Set chkPGP.Container = fraSecurityType(1)
'    '4)
''        fraSecurityPolicy.Width = fraTransactionSecurity.Width '- 2 * fraSecurityPolicy.Left
''        fraSecurityType(1).Width = fraTransactionSecurity.Width - 2 * 12 * Screen.TwipsPerPixelX
''        fraSecurityType(1).Height = fraTransactionSecurity.Height - fraSecurityType(0).Top - 10 * Screen.TwipsPerPixelX
''        fraSecurityType(1).Left = 12 * Screen.TwipsPerPixelX
''        fraSecurityType(1).Caption = "PGP"
''        fraSecurityType(1).ZOrder (0)
''
''        fraPGPOptions.Left = 10 * Screen.TwipsPerPixelX
''        fraPGPOptions.Top = 25 * Screen.TwipsPerPixelX
''
''        cmdBrowseCert(2).Top = fraSecurityType(1).Height - cmdBrowseCert(2).Height - 5 * Screen.TwipsPerPixelX
''        cmdBrowseCert(2).Left = fraPGPOptions.Left
''        cmdBrowseCert(2).Width = 1500
''        cmdBrowseCert(2).Caption = "Assign KeyRings"
''        cmdBrowseCert(2).Visible = True
''
''        'lblPGP(5).Height = cmdBrowseCert(2).Height
''        lblPGP(5).Width = 1000
''        lblPGP(5).Top = cmdBrowseCert(2).Top + 3 * Screen.TwipsPerPixelX
''        lblPGP(5).Left = cmdBrowseCert(2).Left + cmdBrowseCert(2).Width + 10 * Screen.TwipsPerPixelX
''        If OrchestratorXML.getKeyRing(False) <> "" And OrchestratorXML.getKeyRing(True) <> "" Then
''            lblPGP(5).Caption = "Assigned"
''        Else
''            lblPGP(5).Caption = "Not assigned"
''        End If
''        lblPGP(5).Visible = True
''
''    '5)
''        fraSecurityType(1).Visible = True
''    '6)
'        gbPGPDrew = True
'End Sub
'
'Private Sub ShowPGPConventional()
'    hidePGPOptions
''    lblPGP(2).Left = fraSecurityType(1).Width \ 2
'    lblPGP(2).Top = fraPGPOptions.Top + 25 * Screen.TwipsPerPixelX
'    lblPGP(2).Caption = "Password :"
'    lblPGP(2).width = 920
'    lblPGP(2).Visible = True
'
'    chkPGP.Top = lblPGP(2).Top - 2 * Screen.TwipsPerPixelX
'    chkPGP.Left = lblPGP(2).Left + lblPGP(2).width + 600
'    chkPGP.Visible = True
'
'    txtPGP(2).Top = lblPGP(2).Top + lblPGP(2).Height + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).Left = lblPGP(2).Left + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).width = 2650
'    txtPGP(2).Height = 200
'    txtPGP(2).PasswordChar = "*"
'    txtPGP(2).Visible = True
'
'    If Not mrsTransports(xtfINOUT) Then
'        lblPGP(4).Left = lblPGP(2).Left
'        lblPGP(4).Top = txtPGP(2).Top + txtPGP(2).Height + 15 * Screen.TwipsPerPixelX
'        lblPGP(4).Caption = "Algorithm :"
'        lblPGP(4).width = 800
'        lblPGP(4).Visible = True
'
'        cboPGP(0).Left = txtPGP(2).Left
'        cboPGP(0).width = txtPGP(2).width
'        cboPGP(0).Top = lblPGP(4).Top + lblPGP(4).Height + 5 * Screen.TwipsPerPixelX
'
'        cboPGP(0).Clear
'        cboPGP(0).AddItem gstrENCR_3DES
'        cboPGP(0).AddItem gstrENCR_CAST5
'        cboPGP(0).AddItem gstrENCR_IDEA
'        cboPGP(0).AddItem gstrENCR_TWOFISH256
'        cboPGP(0).AddItem gstrENCR_AES128
'        cboPGP(0).AddItem gstrENCR_AES256
'        cboPGP(0).Visible = True
'    End If
'
'End Sub
'
'Private Sub ShowPGPInPublicKey()
'    hidePGPOptions
''    lblPGP(2).Left = fraSecurityType(1).Width \ 2
'    lblPGP(2).Top = fraPGPOptions.Top + 25 * Screen.TwipsPerPixelX
'    lblPGP(2).Caption = "Passphrase :"
'    lblPGP(2).width = 920
'    lblPGP(2).Visible = True
'
'    chkPGP.Top = lblPGP(2).Top - 2 * Screen.TwipsPerPixelX
'    chkPGP.Left = lblPGP(2).Left + lblPGP(2).width + 600
'    chkPGP.Visible = True
'
'    txtPGP(2).Top = lblPGP(2).Top + lblPGP(2).Height + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).Left = lblPGP(2).Left + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).width = 2650
'    txtPGP(2).Height = 1000
'    txtPGP(2).PasswordChar = "*"
'    txtPGP(2).Visible = True
'
'End Sub
'
'
'Private Sub ShowPGPOutPublicKey()
'    hidePGPOptions
'
''    lblPGP(0).Left = fraSecurityType(1).Width \ 2
'    lblPGP(0).Top = fraPGPOptions.Top + fraPGPOptions.Height \ 2 - 10 * Screen.TwipsPerPixelX
'    lblPGP(0).Caption = "Recipient KeyId :"
'    lblPGP(0).width = 1300
'    lblPGP(0).Visible = True
'
'    txtPGP(0).Top = lblPGP(0).Top - 2 * Screen.TwipsPerPixelX
'    txtPGP(0).Left = lblPGP(0).Left + lblPGP(0).width + 5 * Screen.TwipsPerPixelX
'    txtPGP(0).width = 1800
'    txtPGP(0).Height = 200
'    txtPGP(0).Visible = True
'
'    'set the button to view the available keys
'    cmdBrowseCert(4).Top = txtPGP(0).Top
'    cmdBrowseCert(4).Left = txtPGP(0).Left + txtPGP(0).width + 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(4).width = 400 - 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(4).Height = txtPGP(0).Height
'    cmdBrowseCert(4).Caption = "..."
'    cmdBrowseCert(4).Visible = True
'
'End Sub
'
'Private Sub ShowPGPInSignature()
'    hidePGPOptions
''    lblPGP(6).Top = fraSecurityType(1).Height \ 2 - 10 * Screen.TwipsPerPixelX
''    lblPGP(6).Left = fraSecurityType(1).Width \ 2 - 10 * Screen.TwipsPerPixelX
'    lblPGP(6).Caption = "Signature extension : "
'    lblPGP(6).width = 1500
'    lblPGP(6).Visible = True
'
'    txtPGP(3).Top = lblPGP(6).Top - 2 * Screen.TwipsPerPixelX
'    txtPGP(3).Left = lblPGP(6).Left + lblPGP(6).width + 2 * Screen.TwipsPerPixelX
'    txtPGP(3).width = 1800
'    txtPGP(3).Height = 200
'    txtPGP(3).Visible = True
'
'End Sub
'
'Private Sub ShowPGPOutSignature()
'    hidePGPOptions
'
'    lblPGP(6).Top = fraPGPOptions.Top - 15 * Screen.TwipsPerPixelX
''    lblPGP(6).Left = fraSecurityType(1).Width \ 2
'    lblPGP(6).Caption = "Signature extension : "
'    lblPGP(6).width = 1500
'    lblPGP(6).Visible = True
'
'    txtPGP(3).Top = lblPGP(6).Top - 2 * Screen.TwipsPerPixelX
'    txtPGP(3).Left = lblPGP(6).Left + lblPGP(6).width + 2 * Screen.TwipsPerPixelX
'    txtPGP(3).width = 1180
'    txtPGP(3).Height = 200
'    txtPGP(3).Visible = True
'
'
'    lblPGP(1).Left = lblPGP(6).Left ' fraSecurityType(1).Width \ 2
'    lblPGP(1).Top = lblPGP(6).Top + lblPGP(6).Height + 5 * Screen.TwipsPerPixelX
''fraPGPOptions.Top + 15 * Screen.TwipsPerPixelX
'    lblPGP(1).Caption = "Private KeyId :"
'    lblPGP(1).width = 1050
'    lblPGP(1).Visible = True
'
'    txtPGP(1).Top = lblPGP(1).Top - 2 * Screen.TwipsPerPixelX
'    txtPGP(1).Left = lblPGP(1).Left + lblPGP(1).width + 5 * Screen.TwipsPerPixelX
'    txtPGP(1).width = 1600
'    txtPGP(1).Height = 200
'    txtPGP(1).Visible = True
'
'    cmdBrowseCert(3).Top = txtPGP(1).Top
'    cmdBrowseCert(3).Left = txtPGP(1).Left + txtPGP(1).width + 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(3).width = 400 - 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(3).Height = txtPGP(1).Height
'    cmdBrowseCert(3).Caption = "..."
'    cmdBrowseCert(3).Visible = True
'
'    lblPGP(2).Left = lblPGP(1).Left
'    lblPGP(2).Top = lblPGP(1).Top + lblPGP(1).Height + 5 * Screen.TwipsPerPixelX
'    lblPGP(2).Caption = "Passphrase :"
'    lblPGP(2).width = 920
'    lblPGP(2).Visible = True
'
'    chkPGP.Top = lblPGP(2).Top - 2 * Screen.TwipsPerPixelX
'    chkPGP.Left = lblPGP(2).Left + lblPGP(2).width + 600
'    chkPGP.Visible = True
'
'    txtPGP(2).Top = lblPGP(2).Top + lblPGP(2).Height + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).Left = lblPGP(2).Left + 5 * Screen.TwipsPerPixelX
'    txtPGP(2).width = 2650
'    txtPGP(2).Height = 1000
'    txtPGP(2).PasswordChar = "*"
'    txtPGP(2).Visible = True
'
'    lblPGP(3).Left = lblPGP(2).Left
'    lblPGP(3).Top = txtPGP(2).Top + txtPGP(2).Height + 15 * Screen.TwipsPerPixelX
'    lblPGP(3).Caption = "Hash :"
'    lblPGP(3).width = 600
'    lblPGP(3).Visible = True
'
'    cboPGP(1).Top = lblPGP(3).Top - 4 * Screen.TwipsPerPixelX
'    cboPGP(1).Left = lblPGP(3).Left + lblPGP(3).width + 5 * Screen.TwipsPerPixelX
'    cboPGP(1).width = 2100
'
'    cboPGP(1).Clear
'    cboPGP(1).AddItem gstrHASH_SHA1
'    cboPGP(1).AddItem gstrHASH_RIPEMD160
'    cboPGP(1).AddItem gstrHASH_MD5
'    cboPGP(1).Visible = True
'
'
'End Sub
'
'Private Sub ShowPGPInBoth()
'    ShowPGPInPublicKey
'
'    lblPGP(6).Top = txtPGP(2).Top + txtPGP(2).Height + 5 * Screen.TwipsPerPixelX
'    lblPGP(6).Left = lblPGP(2).Left
'    lblPGP(6).Caption = "Signature extension : "
'    lblPGP(6).width = 1500
'    lblPGP(6).Visible = True
'
'    txtPGP(3).Top = lblPGP(6).Top - 2 * Screen.TwipsPerPixelX
'    txtPGP(3).Left = lblPGP(6).Left + lblPGP(6).width + 2 * Screen.TwipsPerPixelX
'    txtPGP(3).width = 1200
'    txtPGP(3).Height = 200
'    txtPGP(3).Visible = True
'
'End Sub
'
'Private Sub ShowPGPOutBoth()
'    ShowPGPOutSignature
'
'    txtPGP(0).Top = cboPGP(1).Top + cboPGP(1).Height + 5 * Screen.TwipsPerPixelX
'    txtPGP(0).Left = txtPGP(1).Left
'    txtPGP(0).width = 1600
'    txtPGP(0).Height = 200
'    txtPGP(0).Visible = True
'
'    cmdBrowseCert(4).Top = txtPGP(0).Top
'    cmdBrowseCert(4).Left = txtPGP(0).Left + txtPGP(0).width + 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(4).width = 400 - 5 * Screen.TwipsPerPixelX
'    cmdBrowseCert(4).Height = txtPGP(0).Height
'    cmdBrowseCert(4).Caption = "..."
'    cmdBrowseCert(4).Visible = True
'
'    lblPGP(0).width = 1300
'    lblPGP(0).Top = txtPGP(0).Top '15 * Screen.TwipsPerPixelX
'    lblPGP(0).Left = txtPGP(0).Left - lblPGP(0).width
'    lblPGP(0).Caption = "Recipient KeyId :"
'    lblPGP(0).Visible = True
'
'End Sub
'
'Private Sub hidePGPOptions()
'        Dim i As Integer
'        For i = 0 To 4
'            lblPGP(i).Visible = False
'        Next i
'        For i = 0 To 3
'            txtPGP(i).Visible = False
'        Next i
'        cboPGP(0).Visible = False
'        cboPGP(1).Visible = False
'        chkPGP.Visible = False
'        cmdBrowseCert(3).Visible = False
'        cmdBrowseCert(4).Visible = False
'        lblPGP(6).Visible = False
'End Sub
'
'Private Sub EnablePGPOptions(bState As Boolean)
'        Dim i As Integer
'        For i = 0 To 4
'            lblPGP(i).Enabled = bState
'        Next i
'        For i = 0 To 3
'            txtPGP(i).Enabled = bState
'        Next i
'        cboPGP(0).Enabled = bState
'        cboPGP(1).Enabled = bState
'        chkPGP.Enabled = bState
'        cmdBrowseCert(3).Enabled = bState
'        cmdBrowseCert(4).Enabled = bState
'        lblPGP(6).Enabled = bState
'        For i = 0 To 3
'            optPGP(i).Enabled = bState
'        Next i
'End Sub


'*****************************************************************
'this sub will load the microsoft saved information about security
'*****************************************************************
Private Sub SetupMSEncControls()

    Dim i As Integer
    Dim enc As EEncoding
    Dim tsymencrypt As TEncryptionSymetric
    Dim astr() As String
    Dim bInSetup As Boolean

    'encryption frame
    cboAlgorithm.Clear
    cboAlgorithm.AddItem gstrENCR_RC2
    cboAlgorithm.AddItem gstrENCR_RC4
    cboAlgorithm.AddItem gstrENCR_DES
    cboAlgorithm.AddItem gstrENCR_3DES

    cboHash.Clear
    cboHash.AddItem gstrHASH_SHA1
    cboHash.AddItem gstrHASH_MD2
    cboHash.AddItem gstrHASH_MD4
    cboHash.AddItem gstrHASH_MD5
    
    'bind apply to
    cboPolicy(1).Clear
    astr = OrchestratorXML.Options_ApplyTo
    For i = 0 To UBound(astr)
        cboPolicy(1).AddItem astr(i)
    Next
    

    'default to pk encryption
    optMSEncryptionType(1).Value = True
    optMSEncryptionType_Click 1
    chkUse(1).Visible = False
    
    chkAttachCert.Value = 0
End Sub


Private Sub LoadPGPInfo()
    Dim tpgpencrypt As TEncryptionPK
    Dim tpgpencryptsym As TEncryptionSymetric
    Dim tpgpsignature As TSignaturePK
    Dim i As Integer
        

    'bug in 1025 3.x new bugs
'    If OrchestratorXML.SecurityType <> esPGP Then
'        OrchestratorXML.RemoveSignature
'        OrchestratorXML.RemoveEncryption
'    End If
'
'
''    txtPGP(0).Text = ""
''    txtPGP(1).Text = ""
''    txtPGP(2).Text = ""
''    cboPGP(0).ListIndex = -1
''    cboPGP(1).ListIndex = -1
''
''    For i = 0 To 3
''        optPGP(i).Value = False
''    Next i
''    hidePGPOptions
''
''    If mrsTransports(xtfINOUT) Then 'incoming
''        optPGP(0).Caption = "Conventional Decryption"
''        optPGP(1).Caption = "Public Key Decryption"
''        optPGP(2).Caption = "Verify Signature"
''        optPGP(3).Caption = "Decrypt and Verify"
''    Else 'outgoing
''        optPGP(0).Caption = "Conventional Encryption"
''        optPGP(1).Caption = "Public Key Encryption"
''        optPGP(2).Caption = "Signature Only"
''        optPGP(3).Caption = "Public Key Encryption + Signature"
''    End If
'    If OrchestratorXML.getKeyRing(False) <> "" And OrchestratorXML.getKeyRing(True) <> "" Then
'        cmdBrowseCert(2).Caption = "Modify Keyrings"
'        lblPGP(5).Caption = "Assigned"
'        EnablePGPOptions True
'    Else
'        cmdBrowseCert(2).Caption = "Assign Keyrings"
'        lblPGP(5).Caption = "Not assigned"
'        EnablePGPOptions False
'    End If
''    chkPGP.Value = vbChecked
''    If Not OrchestratorXML.SecurityType = esPGP Then Exit Sub
'
'    If OrchestratorXML.IsEncrypted And OrchestratorXML.IsSigned Then ' getSignatureType = esPGP Then
'        'both signed and encrypted
'        optPGP(3).Value = True
'        optPGP_Click (3)
'        tpgpencrypt = OrchestratorXML.GetPKEncryption
'        tpgpsignature = OrchestratorXML.getPKSignature
'        If Not mrsTransports(xtfINOUT) Then 'outgoing
'            txtPGP(0).Text = tpgpencrypt.strPublicKey
'            txtPGP(1).Text = tpgpsignature.strPrivateKey
'            txtPGP(2).Text = tpgpsignature.strPassPhrase
'
'            For i = 0 To cboPGP(1).ListCount - 1
'                If OrchestratorXML.Text2Enum_EHashAlgorithm(cboPGP(1).list(i), esPGP) = tpgpsignature.eHash Then
'                    cboPGP(1).ListIndex = i
'                    Exit For
'                End If
'                cboPGP(1).ListIndex = OrchestratorXML.Text2Enum_EHashAlgorithm(gstrHASH_SHA1, esPGP)
'            Next i
''            If CInt(tpgpsignature.eHash) < cboPGP(1).ListCount Then
'                'cboPGP(1).ListIndex = tpgpsignature.eHash
''            End If
'        Else 'incoming
'            txtPGP(2).Text = tpgpencrypt.strPassPhrase
'        End If
'        txtPGP(3).Text = OrchestratorXML.GetSignExtension
'    ElseIf Not OrchestratorXML.IsEncrypted And OrchestratorXML.IsSigned Then
'        'signature only
'        optPGP(2).Value = True
'        optPGP_Click (2)
'        tpgpsignature = OrchestratorXML.getPKSignature
'        If Not mrsTransports(xtfINOUT) Then
'            txtPGP(1).Text = tpgpsignature.strPrivateKey
'            txtPGP(2).Text = tpgpsignature.strPassPhrase
'            'If CInt(tpgpsignature.eHash) < cboPGP(1).ListCount Then
'            For i = 0 To cboPGP(1).ListCount - 1
'                If OrchestratorXML.Text2Enum_EHashAlgorithm(cboPGP(1).list(i), esPGP) = tpgpsignature.eHash Then
'                    cboPGP(1).ListIndex = i
'                    Exit For
'                End If
'                cboPGP(1).ListIndex = OrchestratorXML.Text2Enum_EHashAlgorithm(gstrHASH_SHA1, esPGP)
'            Next i
'                'cboPGP(1).ListIndex = tpgpsignature.eHash
'            'End If
'        End If
'        txtPGP(3).Text = OrchestratorXML.GetSignExtension
'    Else 'encryption only
'        'tpgpencrypt = getPGPEncryption
'        If Not mrsTransports(xtfINOUT) Then 'outgoing
'            If OrchestratorXML.IsSymetric Then
'                'conventional
'                optPGP(0).Value = True
'                optPGP_Click (0)
'                tpgpencryptsym = OrchestratorXML.SymmetricEncryption
'                txtPGP(2).Text = tpgpencryptsym.strPassword
''                If CInt(tpgpencryptsym.eAlgorithm) < cboPGP(0).ListCount Then
''                    cboPGP(0).ListIndex = tpgpencryptsym.eAlgorithm
''                End If
'                For i = 0 To cboPGP(0).ListCount - 1
'                    If OrchestratorXML.Text2Enum_EEncryptionAlgorithm(cboPGP(0).list(i), esPGP) = tpgpencryptsym.eAlgorithm Then
'                        cboPGP(0).ListIndex = i
'                        Exit For
'                    End If
'                    cboPGP(0).ListIndex = OrchestratorXML.Text2Enum_EEncryptionAlgorithm(gstrENCR_3DES, esPGP)
'                Next i
'
'            Else
'                optPGP(1).Value = True
'                optPGP_Click (1)
'                tpgpencrypt = OrchestratorXML.GetPKEncryption
'                txtPGP(0).Text = tpgpencrypt.strPublicKey
'            End If
'        Else    'incoming
'            If OrchestratorXML.IsSymetric Then
'                'conventional
'                optPGP(0).Value = True
'                optPGP_Click (0)
'                tpgpencryptsym = OrchestratorXML.SymmetricEncryption
'                txtPGP(2).Text = tpgpencryptsym.strPassword
''                If CInt(tpgpencrypt.eCipher) < cboPGP(0).ListCount Then
''                    cboPGP(0).ListIndex = CInt(tpgpencrypt.eCipher)
''                End If
'            Else
'                optPGP(1).Value = True
'                optPGP_Click (1)
'                tpgpencrypt = OrchestratorXML.GetPKEncryption
'                'txtPGP(0).Text = tpgpencrypt.strPublicKeyID
'                txtPGP(2).Text = tpgpencrypt.strPassPhrase
'            End If
'        End If
'    End If
End Sub

'*****************************************************************
'this sub will load the microsoft saved information about security
'*****************************************************************
Private Sub SetupMSSigControls()
    Dim i%
    Dim astr() As String

    astr = OrchestratorXML.Options_ApplyTo
    
    cboPolicy(2).Clear
    For i = 0 To UBound(astr)
        cboPolicy(2).AddItem astr(i)
    Next

'    Select Case OrchestratorXML.MessageType
'        Case emtEBXML, emtEBXML_SMIME
'        Case emtSMIME
'            'apply to which part
'            For i = 0 To cboPolicy(2).ListCount - 1
'                If OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(i)) = atCOMPLETE Then
'                    cboPolicy(2).ListIndex = i
'                    Exit For
'                End If
'            Next i
'            cboPolicy(2).Enabled = False
'
'        Case emtCIPHERED
'            'apply to which part
'            For i = 0 To cboPolicy(2).ListCount - 1
'                If OrchestratorXML.Text2Enum_EApplyTo(cboPolicy(2).list(i)) = atCOMPLETE Then
'                    cboPolicy(2).ListIndex = i
'                    Exit For
'                End If
'            Next i
'            cboPolicy(2).Enabled = False
'
'    End Select

    'signature frame
'    txtMSEncryptionType(1).Text = ""
    chkAttachCert.Value = 0
End Sub

Private Sub SecurityControls(EnableDisable As Boolean)
    
    optMSEncryptionType(0).Enabled = EnableDisable
    optMSEncryptionType(1).Enabled = EnableDisable
    optMSEncryptionType(0).Value = EnableDisable
    optMSEncryptionType(1).Value = EnableDisable
    
    lblPloicy(1).Enabled = EnableDisable
    cboPolicy(1).Clear
    cboPolicy(1).Enabled = EnableDisable
    
    lblAlgorithm.Enabled = EnableDisable
    cboAlgorithm.Clear
    cboAlgorithm.Enabled = EnableDisable
    
    lblHash.Enabled = EnableDisable
    cboHash.Clear
    cboHash.Enabled = EnableDisable
    
    
    lblMSEncryptionType(0).Enabled = EnableDisable
    txtMSEncryptionType.Text = ""
    chkUse(1).Enabled = EnableDisable
    
    
    
    lblPloicy(2).Enabled = EnableDisable
    cboPolicy(2).Clear
    cboPolicy(2).Enabled = EnableDisable
    
    
'    btnSigners.Enabled = EnableDisable
    cmdBrowseCert(1).Enabled = EnableDisable
    
    chkAttachCert.Enabled = EnableDisable
End Sub
'*****************************************************************
'this sub will load the microsoft saved information about security
'*****************************************************************
Private Sub LoadMSinfo()

    'set the values depending on the port.xml file
    'encryption on?
    
    
'    Select Case OrchestratorXML.MessageType
'        Case emtEBXML, emtEBXML_SMIME
'            cboPolicy(1).Enabled = True
'            cboPolicy(2).Enabled = True
'        Case emtSMIME, emtCIPHERED
'            cboPolicy(1).Enabled = False
'            cboPolicy(2).Enabled = False
'
'    End Select
    
    
    
'
End Sub


Private Sub chkPGP_Click()
    
    On Error Resume Next
    
    If chkPGP.Value = vbChecked Then
        txtPGP(2).PasswordChar = "*"
    Else
        txtPGP(2).PasswordChar = ""
    End If
End Sub

        

Friend Function FindCertificate_EDI(mediSecurities As Fredi.ediSecurities, strIssuer As String, strID As String, strValidity As String, strSubject As String) As Fredi.ediSecurityCertificate
    Dim strError As String
    Dim strIssuerName As String
    Dim strSerialNumber As String
    Dim strCertSubject As String
    Dim dValidToDate As Date
    
    
    'List all the certificates in a certificate store
    
    Dim oCertLoc As Fredi.ediSecurityCertStoreLocations
    Dim oCertStores As Fredi.ediSecurityCertificateStores
    Dim oCertStore As Fredi.ediSecurityCertificateStore
    Dim ediCert As Fredi.ediSecurityCertificate
    Dim i%, j%, k%
    Dim bFound As Boolean
    
    On Error GoTo ProcessError
    If LenB(strSubject) = 0 Then Exit Function
    Set oCertLoc = mediSecurities.GetCertificateStoreLocations
    Set oCertStores = oCertLoc.GetFirstCertificateStores
    For i = 1 To oCertLoc.Count
        Set oCertStore = oCertStores.GetFirstCertificateStore
        For j = 1 To oCertStores.Count
            Set ediCert = oCertStore.GetFirstCertificate
            For k = 1 To oCertStore.Count
                strIssuerName = ediCert.IssuerName
                strSerialNumber = ediCert.SerialNumber
                dValidToDate = CDate(ediCert.ValidTo)
                strCertSubject = ediCert.SubjectName
                If (strCertSubject = strSubject) Then
                    If (dValidToDate = CDate(strValidity)) Then
                        Set FindCertificate_EDI = ediCert
                        bFound = True
                        Exit For
                    End If
                End If
                Set ediCert = oCertStore.GetNextCertificate
            Next
            If bFound Then Exit For
            Set oCertStore = oCertStores.GetNextCertificateStore
        Next
        If bFound Then Exit For
        Set oCertStores = oCertLoc.GetNextCertificateStores
    Next

    If bFound Then
        mediSecurities.DefaultProviderName = ediCert.CspServiceProvider
        mediSecurities.DefaultCertSystemStoreLocation = oCertStores.location
        mediSecurities.DefaultCertSystemStoreName = oCertStore.StoreName
    Else
        mediSecurities.DefaultProviderName = "Microsoft Strong Cryptographic Provider"
        mediSecurities.DefaultCertSystemStoreLocation = "CurrentUser"
        mediSecurities.DefaultCertSystemStoreName = "My"
    End If

    Set ediCert = Nothing
    Set oCertStore = Nothing
    Set oCertStores = Nothing
    Set oCertLoc = Nothing
    Exit Function
    
ProcessError:
    Set FindCertificate_EDI = Nothing
    Set ediCert = Nothing
    Set oCertStore = Nothing
    Set oCertStores = Nothing
    Set oCertLoc = Nothing

End Function
            
Private Sub cmbCSP_DropDown()

    If cmbCSP.ListCount = 0 Then
    
        Dim oServProvs As Fredi.ediSecurityServiceProviders
        Dim oServProv As Fredi.ediSecurityServiceProvider
        
        Dim i As Integer
        
        Set oServProvs = mediSecurities.GetServiceProviders
        
        Set oServProv = oServProvs.First
        For i = 1 To oServProvs.Count
            cmbCSP.AddItem oServProv.name
            If cmbCSP.Text = oServProv.name Then
                cmbCSP.ListIndex = cmbCSP.ListCount - 1
            End If
            Set oServProv = oServProvs.Next
        Next
    
    End If
    cmbCSP.SetFocus
End Sub
            
Private Sub cmbCertStoreLocation_DropDown()

    If cmbCertStoreLocation.ListCount = 0 Then
        Dim oCertLocs As Fredi.ediSecurityCertStoreLocations
        Dim oCertStores As Fredi.ediSecurityCertificateStores
        Dim oCertStore As Fredi.ediSecurityCertificateStore
        Dim i As Integer
        
        Set oCertLocs = mediSecurities.GetCertificateStoreLocations
    
        Set oCertStores = oCertLocs.GetFirstCertificateStores
        While Not oCertStores Is Nothing
    
            cmbCertStoreLocation.AddItem oCertStores.location
            
            Set oCertStores = oCertLocs.GetNextCertificateStores
        Wend
        
        If Len(Trim(cmbCertStoreLocation.Text)) = 0 Then
            cmbCertStoreLocation.Text = "CurrentUser"
        End If
        
        mediSecurities.DefaultCertSystemStoreLocation = cmbCertStoreLocation.Text

    End If
    
End Sub
        

Private Sub cmbCertStoreName_DropDown()

    If cmbCertStoreName.ListCount = 0 Then

        Dim oCertLocs As Fredi.ediSecurityCertStoreLocations
        Dim oCertStores As Fredi.ediSecurityCertificateStores
        Dim oCertStore As Fredi.ediSecurityCertificateStore
        Dim i As Integer
        
        Set oCertLocs = mediSecurities.GetCertificateStoreLocations
        
        Set oCertStores = oCertLocs.GetCertificateStores(cmbCertStoreLocation.Text)
        Set oCertStore = oCertStores.GetFirstCertificateStore
        For i = 1 To oCertStores.Count
            cmbCertStoreName.AddItem oCertStore.StoreName
            Set oCertStore = oCertStores.GetNextCertificateStore
        Next
        
        If Len(Trim(cmbCertStoreName.Text)) = 0 Then
            cmbCertStoreName.Text = "My"
        End If

        mediSecurities.DefaultCertSystemStoreName = cmbCertStoreName.Text

    End If
    
End Sub
        
Private Sub btnGenerate_Click()
    Dim sServiceProviderName As String
    Dim sCertSubjName As String
    Dim sKeyContainer As String
    Dim sCertificateFileName As String
    Dim sCertIssuerName As String

    On Error GoTo ProcessError
    
    sCertSubjName = txtCertSubjName.Text
'    txtIssuerName.Text = sCertSubjName
'    txtIssuerName.Locked = True
    sCertIssuerName = IIf(txtIssuerName.Text = "", sCertSubjName, txtIssuerName.Text)
    If LenB(Trim(sCertSubjName)) = 0 Then
        Err.Raise 1, "Generate Certificate", "You must have a certificate subject to create it.", _
            vbCritical
    End If
    
    HourGlass Me
    
    sKeyContainer = Trim(Replace(sCertSubjName, " ", "_")) & "_Keys"
    sCertificateFileName = Trim(Replace(sCertSubjName, " ", "")) & ".cer"
    
    'Choose a Cryptographic Service Provider (CSP) that supports encryption and signed algorithms of trading partner
    'For a list of CSP supported by server, use the eSecurity Console
    sServiceProviderName = cmbCSP.Text
    
    ' Set the default service provider name in the securities object so that
    '   we do not have to set it everyhwere else.
    mediSecurities.DefaultProviderName = cmbCSP.Text
    'Create certificate at the following store. (Use eSecurity Console to list store location and names.)
    mediSecurities.DefaultCertSystemStoreLocation = cmbCertStoreLocation.Text
    mediSecurities.DefaultCertSystemStoreName = cmbCertStoreName.Text
    
    
    'Checks to see if certificate exists. Please note that creating a new certificate, even of same name, will
    'creates new public/private key pairs.
    If Not mediSecurities.IsCertificateExists(sCertSubjName) Then
    
        ' Certificate does not exist and so create certificate getting the public key from the
        '   key container specified.
        
        ' The key container that the certificate is created out of must exist.  If it does not already
        '   exist then the key container will be created.
        If CreateTestKeyContainer(mediSecurities, sKeyContainer) Then
        
            Dim sCertificateFile As String
            
            Dim fso As New Scripting.FileSystemObject
            If Not fso.FolderExists(gstrCodePath & "cert") Then
                fso.CreateFolder gstrCodePath & "cert"
            End If
            sCertificateFile = gstrCodePath & "cert\" & sCertificateFileName
            
            'Make certificates valid only for a certain range
            mediSecurities.Property(CertStoreProperty_ValidFrom) = Format$(dpFrom.CurrentDate, "mm/dd/yyyy")
            mediSecurities.Property(CertStoreProperty_ValidTo) = Format$(dpTo.CurrentDate, "mm/dd/yyyy")
            
            ' Create the test certificate.  Get the public key from the AT_KEYEXCHANGE key pair
            '   rather than the AT_SIGNATURE because AT_KEYEXCHANGE can work for both signing
            '   and decrypting while AT_SIGNATURE only works for signing (normally depending on
            '   what algorithm is being used for each key pair per CSP).
            '   The test certificate is self-signed which means that the subject name and the
            '   issuer are the same.
'            mediSecurities.
            If mediSecurities.CreateTestCertificate(sCertificateFile, sCertSubjName, _
                sCertIssuerName, sKeyContainer, CspKeyType_KEYEXCHANGE) <> 1 Then
            
                Err.Raise 1, "Generate Certificate", Err.Description
                
            Else
                'certificate created
                ' Import test certificate.
                Dim oCertificate As Fredi.ediSecurityCertificate
            
                Set oCertificate = mediSecurities.ImportCertificate(sCertificateFile)
                
                Dim oCertLocs As Fredi.ediSecurityCertStoreLocations
                Dim oCertStores As Fredi.ediSecurityCertificateStores
                Dim oCertStore As Fredi.ediSecurityCertificateStore
                
                Set oCertLocs = mediSecurities.GetCertificateStoreLocations
                
                Set oCertStores = oCertLocs.GetCertificateStores(cmbCertStoreLocation.Text)
                Set oCertStore = oCertStores.GetCertificateStore("Root")
                If sCertSubjName = sCertIssuerName Then
                    'add to root store also
                    oCertStore.ImportCertificate sCertificateFile
                End If
                
                ' Update the signer certificate with the key container in the CSP that has the associated
                '   private key used for signing.
                If oCertificate.UpdateCSP(sKeyContainer) <> 1 Then
                    Err.Raise 1, "Generate Certificate", "Colud not update provider with private key." & ";" & Err.Description
                End If
                DisplayStatus "Certificate is created, imported into store, and saved in: " & sCertificateFile
            End If
        Else
            Err.Raise 1, "Generate Certificate", Err.Description
        End If
    Else
        Err.Raise 1, "Generate Certificate", "Certificate already exists in store"
    End If
'    btnImport2Store.Enabled = True
    btnView.Enabled = True
    btnGenerate.Enabled = False

    If gbEncrypted Then
        Dim tencryptEnvelop As TEncryptionEnvelope
        tencryptEnvelop.tCer.strIssuer = oCertificate.IssuerName
        tencryptEnvelop.tCer.strPassPhrase = txtPassphraseCert.Text
        tencryptEnvelop.tCer.strSubject = oCertificate.SubjectName
        OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
    ElseIf gbSigned Then
        Dim tSign() As TCertificate
        tSign(0).strPassPhrase = txtPassphraseCert.Text
        tSign(0).strIssuer = oCertificate.IssuerName
        tSign(0).strSubject = oCertificate.SubjectName
        OrchestratorXML.Signature(gstrLockID) = tSign
    End If

    tbsOptions.TabVisible(7) = True
    tbsOptions.TabVisible(8) = True

'    btnView_Click
    DoEvents
    Sleep 3000
    'save also private key using passphrase
    Set mcapiCert = FindCertificate(oCertificate.SubjectName)
    HourGlass Me
    eplPrivateKey.Properties.Clear
    btnExport_Click
'    If Not mcapiCert Is Nothing Then
'        mcapiCert.Save gstrCodePath & "cert\" & oCertificate.SubjectName & ".pfx", txtPassphraseCert.Text, CAPICOM_CERTIFICATE_SAVE_AS_PFX, CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN
'        DisplayStatus "Certificate with private key is saved in: " & vbCrLf & gstrCodePath & "cert\" & oCertificate.SubjectName & ".pfx"
'    End If
'    tbsOptions_TabChanged
    Exit Sub
ProcessError:
    MsgBox Err.Description, vbCritical, Err.Source
    HideStatus
    HourGlass Me, True
End Sub
        
' CreateTestKeyContainer - Creates a key container in the CSP if the key container
'   does not already exists.
'
'   Input
'   -----
'   oSecurities - Securities object.
'   sContainerName - Key container name.
'
Private Function CreateTestKeyContainer(ByRef oSecurities As Fredi.ediSecurities, ByVal sContainerName As String)

    If Not oSecurities.IsKeyContainerExists(sContainerName) Then
        
        ' Key container does not exist so create key container.
        
        If oSecurities.CreateKeyContainer(sContainerName) = 1 Then
            
            ' Successfully created key container.
            
            CreateTestKeyContainer = True
        Else
        
            CreateTestKeyContainer = False
            
        End If
        
    Else
    
        ' Key container already exists.
        
        CreateTestKeyContainer = True
        
    End If
    
End Function


Friend Function FindCertificate(SubjectName As String) As CAPICOM.ICertificate2
    Dim Cert As New CAPICOM.Certificate
    Dim certs As New CAPICOM.Certificates
    Dim MyStore As New CAPICOM.Store
    Dim i As Integer
    Dim j%
    Dim strError As String
    Dim strSubjectName As String
    Dim strSerialNumber As String
    Dim dValidToDate As Date
    Dim bFound As Boolean
    
    On Error GoTo ProcessError
    For i = 0 To 15
        Select Case i
            Case 0
                MyStore.Open CAPICOM_CURRENT_USER_STORE, _
                        "My", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 1
                MyStore.Open CAPICOM_CURRENT_USER_STORE, _
                        "Root", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 2
                    MyStore.Open CAPICOM_CURRENT_USER_STORE, _
                        "AddressBook", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 3
                    MyStore.Open CAPICOM_CURRENT_USER_STORE, _
                        "CA", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 4
                    MyStore.Open CAPICOM_ACTIVE_DIRECTORY_USER_STORE, _
                        "My", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 5
                    MyStore.Open CAPICOM_ACTIVE_DIRECTORY_USER_STORE, _
                        "Root", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 6
                    MyStore.Open CAPICOM_ACTIVE_DIRECTORY_USER_STORE, _
                        "AddressBook", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 7
                    MyStore.Open CAPICOM_ACTIVE_DIRECTORY_USER_STORE, _
                        "CA", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 8
                    MyStore.Open CAPICOM_LOCAL_MACHINE_STORE, _
                        "My", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 9
                    MyStore.Open CAPICOM_LOCAL_MACHINE_STORE, _
                        "Root", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 10
                    MyStore.Open CAPICOM_LOCAL_MACHINE_STORE, _
                        "AddressBook", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 11
                    MyStore.Open CAPICOM_LOCAL_MACHINE_STORE, _
                        "CA", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 12
                    MyStore.Open CAPICOM_SMART_CARD_USER_STORE, _
                        "My", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 13
                    MyStore.Open CAPICOM_SMART_CARD_USER_STORE, _
                        "Root", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 14
                    MyStore.Open CAPICOM_SMART_CARD_USER_STORE, _
                        "AddressBook", _
                        CAPICOM_STORE_OPEN_READ_ONLY
            Case 15
                    MyStore.Open CAPICOM_SMART_CARD_USER_STORE, _
                        "CA", _
                        CAPICOM_STORE_OPEN_READ_ONLY
        End Select
        
        Set certs = MyStore.Certificates
        
        For Each Cert In certs
            strSubjectName = Cert.GetInfo(CAPICOM_CERT_INFO_SUBJECT_SIMPLE_NAME)
            If (strSubjectName = SubjectName) Then
                Set FindCertificate = Cert
                bFound = True
                Exit For
            End If
        Next
        If bFound Then Exit For
    Next i
    Set Cert = Nothing
    Set certs = Nothing
    Set MyStore = Nothing
    Exit Function
    
ProcessError:
    Set Cert = Nothing
    Set certs = Nothing
    Set MyStore = Nothing
    Set FindCertificate = Nothing

End Function

Private Sub DisplayStatus(Message As String)
    lblInfo.Visible = True
    lblInfo.Text = Message
    lblInfo.Refresh
    Set btnInfo.Picture = img.Picture
    btnInfo.Visible = True
    DoEvents
End Sub

Private Sub HideStatus()
    lblInfo.Visible = False
    Set btnInfo.Picture = Nothing
    btnInfo.Visible = False
End Sub

Private Sub txtPassphrase_Change()
    
    On Error Resume Next

    OrchestratorXML.PassPhraseSSH(gstrLockID) = txtPassphrase.Text


End Sub

Private Sub txtPassphraseCert_Change()
    If mbInSetup Then Exit Sub
    If (LenB(txtPassphraseCert.Text) <> 0) And (LenB(txtCertSubjName.Text) <> 0) Then
        btnExport.Enabled = True
        If gbEncrypted Then
            Dim tencryptEnvelop As TEncryptionEnvelope
            tencryptEnvelop = OrchestratorXML.EnvEncryption(gstrLockID)
            tencryptEnvelop.tCer.strPassPhrase = txtPassphraseCert.Text
            OrchestratorXML.EnvEncryption(gstrLockID) = tencryptEnvelop
        ElseIf gbSigned Then
            Dim tSign() As TCertificate
            tSign = OrchestratorXML.Signature(gstrLockID)
            tSign(0).strPassPhrase = txtPassphraseCert.Text
            OrchestratorXML.Signature(gstrLockID) = tSign
        End If
    End If

End Sub

