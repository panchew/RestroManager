VERSION 5.00
Object = "{13677477-3E5D-4A74-8C0C-C63F18FA06EF}#1.0#0"; "vbsini.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Acceso al sistema"
   ClientHeight    =   2460
   ClientLeft      =   1890
   ClientTop       =   1845
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4665
   Begin CtlIni.VBSIni ocxINI 
      Left            =   2160
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtContraseña 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdClave 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4140
      TabIndex        =   2
      Top             =   1320
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cdgFileOpen 
      Left            =   720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1515
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClave_Click()
      If txtUsuario.Text <> "" Then
            VerificarUsuario
      Else
            MsgBox "Debe especificar un nombre de usuario"
      End If
End Sub

Private Sub Form_Load()
  OpenDatabase
End Sub
