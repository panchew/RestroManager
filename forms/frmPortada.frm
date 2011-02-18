VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenuOpciones 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RestroManager"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   18915
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   18915
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   11460
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   10515
      Begin VB.Image imgReportes 
         Height          =   2055
         Left            =   4200
         Picture         =   "frmMenuOpciones.frx":0442
         Stretch         =   -1  'True
         Top             =   7500
         Width           =   1995
      End
      Begin VB.Image imgVentas 
         Height          =   2055
         Left            =   240
         Picture         =   "frmMenuOpciones.frx":CC8C
         Stretch         =   -1  'True
         Top             =   7500
         Width           =   1995
      End
      Begin VB.Image imgAlmacen 
         Height          =   2055
         Left            =   2280
         Picture         =   "frmMenuOpciones.frx":194D6
         Stretch         =   -1  'True
         Top             =   7500
         Width           =   1965
      End
      Begin VB.Label lblVentas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "VENTAS"
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   9540
         Width           =   1875
      End
      Begin VB.Label lblAlmacen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "ALMACEN"
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   9540
         Width           =   1935
      End
      Begin VB.Label lblReportes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "REPORTES"
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   9540
         Width           =   1815
      End
      Begin VB.Label lblAdmin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "ADMINISTRACION"
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   9540
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "SALIR"
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   8280
         TabIndex        =   1
         Top             =   9540
         Width           =   1875
      End
      Begin VB.Image imgAdministracion 
         Height          =   2055
         Left            =   6300
         Picture         =   "frmMenuOpciones.frx":25D20
         Stretch         =   -1  'True
         Top             =   7500
         Width           =   1995
      End
      Begin VB.Image imgSalir 
         Height          =   1920
         Left            =   8220
         Picture         =   "frmMenuOpciones.frx":3256A
         Stretch         =   -1  'True
         Top             =   7500
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2115
         Index           =   4
         Left            =   8220
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2115
         Index           =   3
         Left            =   6225
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2115
         Index           =   2
         Left            =   4230
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2115
         Index           =   1
         Left            =   2235
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   2115
         Index           =   0
         Left            =   240
         Top             =   7800
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   7800
         Left            =   240
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10005
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMenuOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub cmdVentas_Click()
      frmCroquis.Show
End Sub

Private Sub Form_DblClick()
      CommonDialog1.ShowColor
      Me.BackColor = CommonDialog1.Color
      MsgBox Hex(CommonDialog1.Color)
End Sub

Private Sub Form_Load()
      Dim intLinea As Integer
      
      Me.Left = 0
      Me.Top = 0
      Me.Width = Screen.Width
      Me.Height = Screen.Height '- 500
      Frame1.Left = (Me.Width - Frame1.Width) / 2

      strSQL = "SELECT * FROM Empresa"
      rs.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
      If rs.EOF = False Then
          strEmpresa = rs!Empresa
          Image1.Picture = LoadPicture(strRuta & rs!ImagenMenuOpciones)
          Me.Caption = Me.Caption & " - " & rs!Empresa
          'Me.BackColor = Hex$(rs!ColorFondo)
      Else
          strEmpresa = ""
      End If
                            
End Sub

Private Sub Form_Unload(Cancel As Integer)
      End
End Sub

Private Sub imgAdministracion_Click()
      frmCatalogoProductos.Show
End Sub

Private Sub imgReportes_Click()
      frmReporteTotales.Show
End Sub

Private Sub imgSalir_Click()
      Unload Me
End Sub

Private Sub imgVentas_Click()
      frmCroquis.Show
End Sub

Private Sub Label1_Click()
      Unload Me
End Sub

Private Sub lblAdmin_Click()
      frmCatalogoProductos.Show
End Sub

Private Sub lblReportes_Click()
      frmReporteTotales.Show
End Sub

Private Sub lblVentas_Click()
      frmCroquis.Show
End Sub
