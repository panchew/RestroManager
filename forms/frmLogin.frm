VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{13677477-3E5D-4A74-8C0C-C63F18FA06EF}#1.0#0"; "vbsini.ocx"
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
      TabIndex        =   3
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
      TabIndex        =   0
      Top             =   1320
      Width           =   435
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
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
      TabIndex        =   2
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

Private Sub VerificarUsuario()
      Dim rs As New ADODB.Recordset

      strSQL = "SELECT * FROM tblUsuarios WHERE Nombre = '" & txtUsuario.Text & "'"
      rs.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
      If rs.EOF = False Then
            If rs!Password = txtContraseña.Text Then
                  strIdUsuario = rs!IdUsuario
                  ObtenerCorteActual
                  frmMenuOpciones.Show
                  Unload Me
            Else
                  MsgBox "Contraseña incorrecta"
            End If
      Else
          MsgBox "Usuario no econtrado"
      End If

End Sub

Private Sub ObtenerCorteActual()
      Dim rs As New ADODB.Recordset
      
      strSQL = "SELECT * FROM tblCorteCaja WHERE FechaCorte is NULL"
      rs.Open strSQL, connADO, adOpenKeyset, adLockOptimistic, adCmdText
      If rs.EOF = False Then
            strIdCorteActual = rs!IdCorte
      End If
      rs.Close
      Set rs = Nothing

End Sub

Private Sub Form_Load()

On Error GoTo ErrError
      'Obtener información de archivo de almacenamiento de datos
      ocxINI.Archivo = App.Path & "\options.ini"
      ocxINI.Seccion = "DATOS"
      ocxINI.LLave = "RUTA"
      strRuta = ocxINI.LeeIni
      ocxINI.LLave = "ARCHIVO"
      strArchivo = ocxINI.LeeIni
      MsgBox strRuta & strArchivo
      strRutaConn = strRuta & strArchivo
      
Abrir:
      
      If strRutaConn <> "" Then
            connADO.Provider = "Microsoft.Jet.OLEDB.4.0"
            connADO.Properties("Data Source") = strRutaConn
            connADO.Properties("Jet OLEDB:Database Password") = "restro"
            connADO.Open
            commADO.ActiveConnection = connADO
            GoTo Fin
      Else
            Err.Number = 32755
            GoTo ErrError
      End If
        
ErrError:
    Select Case Err.Number
        Case 0
            GoTo Fin
        Case -2147467259
            'MsgBox "No se ha encontrado el archivo de base de datos; favor de elegir uno", vbCritical + vbOKOnly, "Error"
            CommonDialog1.DefaultExt = "mdb"
            CommonDialog1.Filter = "Base de datos|*.mdb; Todos los archivos|*.*"
            CommonDialog1.ShowOpen
            strRutaConn = CommonDialog1.FileName
            GoTo Abrir
        
        Case 32755, -2147217843
            MsgBox "No se ha elegido ningún archivo de datos; las opciones no estarán disponibles", vbCritical + vbOKOnly, "Error"
            Exit Sub
        Case Else
            MsgBox "Se ha presentado el siguiente error: " & vbCrLf & "(" & Err.Number & ")" & Err.Description
    End Select
    
Fin:
      
End Sub
